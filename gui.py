import tkinter as tk
import attributes as att
import main
import sys
import tkinter.messagebox as msgbox


def run_main_task():
    try:
        main.main()
        msgbox.showinfo("Notification", "Task Completed Successfully!")
        sys.exit()
    except Exception as e:
        msgbox.showerror("Error", f"An error occurred: {e}")
        sys.exit()


main_window = tk.Tk()
main_window.title('4G Automation__mzfarisi@gmail.com')
main_window.geometry('600x400')
main_window.iconbitmap(
    "D:/MINE/PYTHON/PYTHON_PROJECTS/excel_automation/_4G/4G_automatic_report/import-from-excel-512.ico")
run_button = tk.Button(main_window, text='Run', command=run_main_task)
exit_button = tk.Button(main_window, text='Exit', command=sys.exit)

run_button.pack()
exit_button.pack(expand=1)

main_window.mainloop()
