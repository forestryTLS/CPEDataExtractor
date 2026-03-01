from utils.common import (
    ENROLMENTS_DATE_PRESET_KEY,
    ENROLMENTS_DATE_FROM_KEY,
    ENROLMENTS_DATE_TO_KEY,

    USERS_DATE_PRESET_KEY,
    USERS_DATE_FROM_KEY,
    USERS_DATE_TO_KEY
)

def display_filter_config_gui(
    enrollments_settings: dict,
    users_settings: dict
):
    import tkinter as tk
    from tkinter import ttk
    from tkcalendar import DateEntry

    enrollments_settings['filter'][ENROLMENTS_DATE_PRESET_KEY] = 'custom'
    users_settings['filter'][USERS_DATE_PRESET_KEY] = 'custom'

    enrolments_initial_from = enrollments_settings['filter']['enrollment_date_from']
    enrolments_initial_to = enrollments_settings['filter']['enrollment_date_to']

    # initialize the app and create a frame to hold widgets
    root = tk.Tk()
    frm = ttk.Frame(root, padding=10)
    frm.grid()

    # define labels and date widgets to filter enrolments and users by
    ttk.Label(frm, text='FILTER RECORDS').grid(row=0, column=0, columnspan=5)
    ttk.Label(frm, text="From:").grid(row=1, column=0)

    sv_date_from = tk.StringVar()
    sv_date_to = tk.StringVar()

    start_date_entry = DateEntry(
        frm, 
        date_pattern='yyyy-MM-dd',
        textvariable=sv_date_from
    )
    
    start_date_entry.grid(row=1, column=1)

    sv_date_from.set(enrolments_initial_from)

    ttk.Label(frm, text='-').grid(row=1, column=2)

    ttk.Label(frm, text="To:").grid(row=1, column=3)

    end_date_entry = DateEntry(
        frm,
        date_pattern='yyyy-MM-dd',
        textvariable=sv_date_to
    )
    
    end_date_entry.grid(row=1, column=4)

    sv_date_to.set(enrolments_initial_to)

    def set_enrolments_date_from(sv, index, mode):
        from_date = sv_date_from.get()

        enrollments_settings['filter'][ENROLMENTS_DATE_FROM_KEY] = from_date 
        users_settings['filter'][USERS_DATE_FROM_KEY] = from_date

    def set_enrolments_date_to(sv, index, mode):
        to_date = sv_date_to.get()

        enrollments_settings['filter'][ENROLMENTS_DATE_TO_KEY] = to_date
        users_settings['filter'][USERS_DATE_TO_KEY] = to_date

    sv_date_from.trace_add('write', set_enrolments_date_from)
    sv_date_to.trace_add('write', set_enrolments_date_to)

    ttk.Button(frm, text="Save", command=root.destroy).grid(row=2, column=0, columnspan=5)
    root.mainloop()

    return enrollments_settings, users_settings