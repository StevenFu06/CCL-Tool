from compare import Rearrange, Bom, Tracker
from docx.api import Document
from docx.shared import RGBColor
from package import Parser, Parent
import pandas as pd
import json


class CCL:

    def __init__(self, word_doc, avl_bom):
        self.document = Document(word_doc)
        self.avl_bom = pd.read_csv(avl_bom)
        self.avl_bom_path = avl_bom
        self.table = self.document.tables[0]

    def update(self, new_avl_bom_path, tree_old=None, tree_new=None):
        # NOte to Steven, NEED A WAY TO KEEP TRACK OF DUPLICATES!!!!

        if tree_old is None:
            tree_old = Parent(self.avl_bom_path).build_tree()
        if tree_new is None:
            tree_new = Parent(new_avl_bom_path).build_tree()

        # Create Bom Objects and tracker objects
        new_avl_bom = pd.read_csv(new_avl_bom_path)
        bom_old = Bom(self.avl_bom, tree_old)
        bom_new = Bom(new_avl_bom, tree_new)
        track = Tracker()

        Rearrange(bom_old, bom_new, track)
        all_updated = track.combine_found()
        for row in self.table.rows:
            if row.cells[0].text in all_updated['old_pn'].to_list():
                update_info = all_updated.loc[all_updated['old_pn'] == row.cells[0].text]
                new_index = update_info['new_idx'].values[0]

                # Change font color
                for i in range(4):
                    font = row.cells[i].paragraphs[0].runs[0].font
                    font.color.rgb = RGBColor(255, 0, 0)

                ## update pn ##
                row.cells[0].text = update_info['new_pn'].values[0]

                ## update desc + fn ##
                desc = new_avl_bom.loc[new_index, 'Description']
                fn = new_avl_bom.loc[new_index, 'F/N']
                row.cells[1].text = f'{desc} (#{fn})'

                ## update manufacturer ##
                # Pop manufacturers one at a time because multiple can exist
                to_pop = new_avl_bom.loc[new_index, 'Manufacturer'].split('\n')[0]
                new_avl_bom.loc[new_index, 'Manufacturer'] = \
                    new_avl_bom.loc[new_index, 'Manufacturer'].replace(to_pop + '\n', '')
                # if empty replace with AB sciex but also warn the user
                if to_pop.isspace():
                    to_pop = 'AB Sciex'
                    print(f'{update_info["new_pn"].values[0]} couldnt find manufacturer')
                row.cells[2].text = to_pop

                ## Update Equivalent or type/model ##
                # Pop Equivalent one at a time because multiple can exist
                to_pop = new_avl_bom.loc[new_index, 'Equivalent'].split('\n')[0]
                new_avl_bom.loc[new_index, 'Equivalent'] = \
                    new_avl_bom.loc[new_index, 'Equivalent'].replace(to_pop + '\n', '')
                # if empty replace with part number only but also warn the user
                if to_pop.isspace():
                    to_pop = update_info['new_pn'].values[0]
                    print(f'{update_info["new_pn"].values[0]} couldnt find Equivalent')

                row.cells[3].text = to_pop

    def save(self):
        self.document.save('test.docx')


if __name__ == '__main__':
    table = Parser('Fake Bom\\Sample CCL with enter.docx')

    with open('treea.json', 'r') as read:
        old_tree = json.load(read)
    old_bom = pd.read_csv('reva.csv')
    bom_old = Bom(old_bom, old_tree)

    with open('treec.json', 'r') as read:
        new_tree = json.load(read)
    new_bom = pd.read_csv('revc.csv')
    bom_new = Bom(new_bom, new_tree)

    tracker = Tracker()
    Rearrange(bom_old, bom_new, tracker)
    print(tracker.combine_found())
    # ccl = CCL('rev a bugatti.docx', 'reva.csv')
    # ccl.update('revc.csv', old_tree, new_tree)
    # # ccl.save()
