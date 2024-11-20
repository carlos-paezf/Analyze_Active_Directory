from modules import App_GUI_Manager, Group_Manager, Excel_Manager


if __name__ == "__main__":
    group_manager = Group_Manager()
    excel_manager = Excel_Manager()
    app_ui_manager = App_GUI_Manager(group_manager, excel_manager)

    app_ui_manager.mainloop()