from docx.api import Document
from docx.shared import RGBColor
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import Pt

import pandas as pd
import json
import os

from enovia import Enovia
from package import Parser, Parent, _re_doc_num, _re_pn
from compare import Rearrange, Bom, Tracker
from filehandler import *

class CCL:
    def __init__(self):
        # Files
        self.ccl_docx = None  # Docx path
        self.filtered = None  # df
        self.avl_bom = None  # df
        self.avl_bom_updated = None  # df
        # Save paths
        self.avl_bom_path = None
        self.avl_bom_updated_path = None
        self.path_illustration = None
        self.path_ccl_data = None
        self.path_checks = []
        # Enovia
        self.username = None
        self.password = None

########################################################################################################################
# Bom comparison
########################################################################################################################

    def set_bom_compare(self, avl_bom_old, avl_bom_new):
        self.avl_bom_path = avl_bom_old
        self.avl_bom_updated_path = avl_bom_new

        self.avl_bom = pd.read_csv(avl_bom_old)
        self.avl_bom_updated = pd.read_csv(avl_bom_new)

    def avl_path_to_df(self):
        self.avl_bom = pd.read_csv(self.avl_bom_path)
        self.avl_bom_updated = pd.read_csv(self.avl_bom_updated_path)

    def bom_compare(self):
        # ISSUE where bom_compare outputs nothing, needs to be converted to df or something
        if self.avl_bom is None or self.avl_bom_updated is None:
            raise ValueError('Missing required fields, ccl, avl_new or avl_old')

        tree, bom, tree_updated, bom_updated = self._get_bom_obj()
        tracker = Tracker()
        Rearrange(bom, bom_updated, tracker)

        tree, bom, tree_updated, bom_updated = self._get_bom_obj()
        tracker_reversed = Tracker()
        Rearrange(bom_updated, bom, tracker_reversed)

        return tracker, tracker_reversed

    def _get_bom_obj(self):
        tree = Parent(self.avl_bom_path).build_tree()
        bom = Bom(pd.read_csv(self.avl_bom_path), tree)

        tree_updated = Parent(self.avl_bom_updated_path).build_tree()
        bom_updated = Bom(pd.read_csv(self.avl_bom_updated_path), tree_updated)
        return tree, bom, tree_updated, bom_updated

########################################################################################################################
# CCL Updating
########################################################################################################################

    def update_ccl(self, save_path, ccl_docx=None):
        if ccl_docx is not None:
            self.ccl_docx = ccl_docx
        elif self.ccl_docx.docx is None:
            raise ValueError('CCL is not given')

        tracker = self.bom_compare()[0]
        all_updates, removed = tracker.combine_found(), tracker.not_found_to_df()
        ccledit = CCLEditor(self.ccl_docx)
        self._updates_only(ccledit, all_updates)

        ccledit.save(save_path)

    def _updates_only(self, ccledit, all_updates):
        for row in range(len(ccledit.table.rows)):
            pn = ccledit.get_text(row, 0)
            if pn in all_updates['old_pn'].to_list():
                to_update = all_updates.loc[pn == all_updates['old_pn']].values[0]
                # previous formatting
                bold = ccledit.isbold(row, 0)
                # Update fields
                self._update_pn(row, to_update, ccledit)
                self._update_desc_fn(row, to_update, ccledit)
                self._update_manufacturer(row, to_update, ccledit)
                self._update_model(row, to_update, ccledit)
                if bold:
                    ccledit.bold_row(row)
                self._match_conditions(row, ccledit, to_update)
                # Remove updated from df
                all_updates = all_updates[all_updates.old_idx != to_update[0]]

    def _update_pn(self, row, to_update, ccledit):
        align = ccledit.get_justification(row, 0)
        ccledit.set_text(row, 0, to_update[3])
        ccledit.set_justification(row, 0, align)

    def _update_desc_fn(self, row, to_update, ccledit):
        desc = self.avl_bom_updated.loc[to_update[2], 'Description']
        fn = self.avl_bom_updated.loc[to_update[2], 'F/N']
        ccledit.set_text(row, 1, f'{desc} (#{fn})')

    def _update_manufacturer(self, row, to_update, ccledit):
        # Pop manufacturers one at a time because multiple can exist
        manufacturer = self.avl_bom_updated.loc[to_update[2], 'Manufacturer'].split('\n')[0]
        self.avl_bom_updated.loc[to_update[2], 'Manufacturer'] = \
            self.avl_bom_updated.loc[to_update[2], 'Manufacturer'].replace(manufacturer + '\n', '')
        # if empty replace with AB sciex but also warn the user
        if manufacturer.isspace():
            manufacturer = 'AB Sciex'
            print(f'{to_update[3]} couldnt find manufacturer')
        ccledit.set_text(row, 2, manufacturer)

    def _update_model(self, row, to_update, ccledit):
        # Pop Equivalent one at a time because multiple can exist
        model = self.avl_bom_updated.loc[to_update[2], 'Equivalent'].split('\n')[0]
        self.avl_bom_updated.loc[to_update[2], 'Equivalent'] = \
            self.avl_bom_updated.loc[to_update[2], 'Equivalent'].replace(model + '\n', '')
        # if empty replace with part number only but also warn the user
        if model.isspace():
            model = to_update[3]
            print(f'{to_update[3]} couldnt find Equivalent')
        ccledit.set_text(row, 3, model)

    def _match_conditions(self, row, ccledit, to_update):
        if to_update[4] == 'full':
            ccledit.highglight_row(row, 'YELLOW')
        elif to_update[4] == 'partial':
            ccledit.highglight_row(row, 'RED')
        elif to_update[4] == 'fn_only':
            ccledit.highglight_row(row, 'YELLOW')

    def _removed_only(self, ccledit, removed):
        for row in range(len(ccledit.table.rows)):
            pn = ccledit.get_text(row, 0)
            if pn in removed['pn'].to_list():
                ccledit.strike_row(row)

########################################################################################################################
# Specification Documents Gathering
########################################################################################################################

    def collect_documents(self, processes=1, headless=True):
        if self.ccl_docx is None:
            raise ValueError('CCL document is not given')
        if self.username is None or self.password is None:
            raise ValueError('Enovia username or password not given')
        if self.path_ccl_data is None:
            raise ValueError('CCL Documents save location not given')

        collector = DocumentCollector(username=self.username,
                                      password=self.password,
                                      ccl=self.ccl_docx,
                                      save_dir=self.path_ccl_data,
                                      processes=processes,
                                      headless=headless)
        collector.collect_documents(check_paths=self.path_checks)

########################################################################################################################
# Illustration gathering and CCL updating
########################################################################################################################

    def collect_illustrations(self, processes=1):
        if self.ccl_docx is None:
            raise ValueError('CCL document is not given')
        if self.path_illustration is None:
            raise ValueError('Illustration save path not given')
        illustration = Illustration(ccl=self.ccl_docx,
                                    save_dir=self.path_illustration,
                                    processes=processes)
        illustration.get_illustrations(ccl_dir=self.path_ccl_data)

    def insert_illustration_data(self, save_path):
        if self.ccl_docx is None:
            raise ValueError('CCL Document is not given')
        ccledit = CCLEditor(self.ccl_docx)
        for row in range(len(ccledit.table.rows)):
            pn = _re_pn(ccledit.get_text(row, 0))
            bold = ccledit.isbold(row, 4)

            new_technical = self.new_illustration_data(pn) + self.remove_illustration_data(ccledit.get_text(row, 4))
            ccledit.set_text(row, 4, new_technical)

            if bold:
                ccledit.set_bold(row, 4)
        ccledit.save(save_path)

    def new_illustration_data(self, pn):
        info = []
        for file in os.listdir(self.path_illustration):
            if file.endswith('.pdf'):
                ill_num = file.split(' ')[0]
                file_pn = file.split(' ')[1]
                dnum = _re_doc_num(file)
                sch_assy = 'Assy.' if 'Assy.' in file.split(' ') else 'Sch.'
                if str(pn) == str(file_pn):
                    info.append((ill_num, dnum[0], sch_assy))

        illustration_data = 'Refer to'
        if info:
            for ill_num, dnum, sch_assy in info:
                illustration_data = illustration_data + f' {ill_num} {sch_assy} {dnum};'
            return illustration_data
        return ''

    @staticmethod
    def remove_illustration_data(technical_string):
        results = re.findall(r'(?:\s*,|and)?\s*(?:Refer to)?\s*(?:Ill.|Ill)\s*(?:\d+.|\d+)\s*'
                             r'(?:Sch.|Assy.|Sch|Assy)\s*D\d+\s*(?:;|and)?',
                             technical_string, re.IGNORECASE)
        for result in results:
            technical_string = technical_string.replace(result, '')
        return technical_string

    def insert_illustration(self, ill_num, new_ill, save_path):
        illustration = Illustration(self.ccl_docx, self.path_illustration)
        illustration.shift_up_ill(ill_num)
        # DEAL WITH PATH MANAGEMENT WITH TKINTER DONT FORGET
        copyfile(new_ill, self.path_illustration)
        self.insert_illustration_data(save_path)

    def delete_illustration(self, ill_num, rm_ill, save_path):
        illustration = Illustration(self.ccl_docx, self.path_illustration)
        illustration.shift_down_ill(ill_num)
        os.remove(rm_ill)
        # DEAL WITH PATH MANAGEMENT WITH TKINTER DONT FORGET
        self.insert_illustration_data(save_path)


class CCLEditor:

    def __init__(self, docx_path):
        self.document = Document(docx_path)
        self.table = self.document.tables[0]

    def get_text(self, row, column):
        return self.table.rows[row].cells[column].text

    def set_text(self, row, column, new_string):
        self.table.rows[row].cells[column].text = new_string

    def set_bold(self, row, column):
        for paragraph in self.table.rows[row].cells[column].paragraphs:
            for run in paragraph.runs:
                run.font.bold = True

    def bold_row(self, row):
        for cell in self.table.rows[row].cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True

    def isbold(self, row, column):
        for paragraph in self.table.rows[row].cells[column].paragraphs:
            for run in paragraph.runs:
                if run.font.bold:
                    return True
        return False

    def set_italic(self, row, column):
        for paragraph in self.table.rows[row].cells[column].paragraphs:
            for run in paragraph.runs:
                run.font.italic = True

    def italic_row(self, row):
        for cell in self.table.rows[row].cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.italic = True

    def isitalic(self, row, column):
        for paragraph in self.table.rows[row].cells[column].paragraphs:
            for run in paragraph.runs:
                if run.font.italic:
                    return True
        return False

    def set_highlight(self, row, column, colour):
        for paragraph in self.table.rows[row].cells[column].paragraphs:
            for run in paragraph.runs:
                run.font.highlight_color = getattr(WD_COLOR_INDEX, colour)

    def highlight_row(self, row, colour):
        for cell in self.table.rows[row].cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.italic = getattr(WD_COLOR_INDEX, colour)

    def set_colour(self, row, column, r, g, b):
        for paragraph in self.table.rows[row].cells[column].paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(r, g, b)

    def colour_row(self, row, r, g, b):
        for cell in self.table.rows[row].cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.color.rgb = RGBColor(r, g, b)

    def get_font(self, row, column):
        return self.table.rows[row].cells[column].paragraphs[0].runs[0].font.name

    def set_font(self, row, column, font):
        for paragraph in self.table.rows[row].cells[column].paragraphs:
            for run in paragraph.runs:
                run.font.name = font

    def font_row(self, row, font):
        for cell in self.table.rows[row].cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = font

    def set_fontsize(self, row, column, size):
        for paragraph in self.table.rows[row].cells[column].paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(size)

    def fontsize_row(self, row, size):
        for cell in self.table.rows[row].cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(size)

    def get_justification(self, row, column):
        return self.table.rows[row].cells[column].paragraphs[0].alignment

    def set_justification(self, row, column, justification):
        # just = {'left': 0, 'center': 1, 'right': 3, 'distribute': 4}
        for paragraph in self.table.rows[row].cells[column].paragraphs:
            paragraph.alignment = justification

    def set_strike(self, row, column):
        for paragraph in self.table.rows[row].cells[column].paragraphs:
            for run in paragraph.runs:
                run.font.strike = True

    def strike_row(self, row):
        for cell in self.table.rows[row].cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.strike = True

    def save(self, path):
        self.document.save(path)


if __name__ == '__main__':
    import pandas as pd
    ccl = CCL()
    ccl.ccl_docx = 'ccl.docx'
    # ccl.set_bom_compare('wombat revg.csv', 'wombat revn.csv')
    # ccl.update_ccl('test.docx')
    ccl.path_illustration = 'Annex A - Illustrations'
    ccl.insert_illustration_data('test.docx')
