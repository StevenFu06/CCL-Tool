global total
global current

def add_total(val):
    global total
    total += val

def add_current(val):
    global current
    current += val

def current_progress():
    global total, current
    return current/total

def reset():
    global total, current
    total = 0
    current = 0
