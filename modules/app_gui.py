import tkinter as tk
from tkinter import messagebox
from modules import Group_Manager, Excel_Manager, Search_Active_Directory

class App_GUI_Manager(tk.Tk):
    def __init__(self, group_manager: Group_Manager, excel_manager: Excel_Manager):
        super().__init__()

        self.group_manager = group_manager
        self.excel_manager = excel_manager

        self.title = "Gestión de Grupos en Directorio Activo"

        self.geometry("1200x900")

        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        self.create_widgets()


    
    def create_widgets(self):
        """
        The function `create_widgets` creates a listbox, label, entry field, and buttons for adding and
        removing groups in a GUI interface.
        """
        self.groups_frame = tk.Frame(self)
        self.groups_frame.pack(pady=10)

        tk.Label(self.groups_frame, text="Nombre del grupo:").pack(side=tk.LEFT, padx=5)
        
        self.entry_group = tk.Entry(self.groups_frame)
        self.entry_group.pack(side=tk.LEFT, padx=5)
        add_btn = tk.Button(self.groups_frame, text="Agregar", command=self.add_group)
        add_btn.pack(side=tk.LEFT, padx=5)


        self.listbox_groups = tk.Listbox(self, width=50, height=40)
        self.listbox_groups.pack(pady=25)
        remove_btn = tk.Button(self.groups_frame, text="Eliminar", command=self.remove_group)
        remove_btn.pack(side=tk.LEFT, padx=5)

        self.load_groups_to_listbox()

        self.options_frame = tk.Frame(self)
        self.options_frame.pack(pady=10)
        
        report_btn = tk.Button(self.options_frame, text="Generar Reporte Excel", command=self.generate_excel_report)
        report_btn.pack(side=tk.LEFT, padx=5)



    
    def load_groups_to_listbox(self):
        """
        This function clears a listbox and populates it with groups loaded from a file using a group
        manager.
        """
        self.listbox_groups.delete(0, tk.END)
        self.groups = self.group_manager.read_groups_from_file()

        for group in self.groups:
            self.listbox_groups.insert(tk.END, group)

    
    def add_group(self):
        """
        The `add_group` function adds a new group, updates the listbox, and displays a success or error
        message using tkinter messagebox.
        """
        group_name = self.entry_group.get().strip()
        success, message = self.group_manager.add_group(group_name)
        
        if success:
            self.load_groups_to_listbox()
            self.entry_group.delete(0, tk.END)
            messagebox.showinfo("Éxito", message)
        else:
            messagebox.showerror("Error", message)


    def remove_group(self):
        """
        This Python function removes a selected group from a listbox and displays success or error
        messages accordingly.
        """
        selected_group = self.listbox_groups.get(tk.ACTIVE)
        
        if selected_group:
            success, message = self.group_manager.remove_group(selected_group)
        
            if success:
                self.load_groups_to_listbox()
                messagebox.showinfo("Éxito", message)
            else:
                messagebox.showinfo("Error", message)
        else:
            messagebox.showerror("Error", "Seleccionar un grupo a remover")


    def generate_excel_report(self):
        success, message = self.excel_manager.save_to_excel(
            group_members=Search_Active_Directory(self.groups).group_members
        )

        if success:
            messagebox.showinfo("Éxito", message)
        else:
            messagebox.showerror("Error", "Revisa la consola")