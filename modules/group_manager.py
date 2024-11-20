from pathlib import Path

from measure_run_time import measure_run_time

class Group_Manager():
    def __init__(self):
        self.parameters_path = Path("./parameters.txt")


    @measure_run_time
    def read_groups_from_file(self) -> list[str]:
        """
        This Python function reads and returns a list of groups from a file specified by the
        `parameters_path`.
        :return: A list of strings is being returned.
        """
        if not self.parameters_path.exists():
            raise FileNotFoundError("Archivo de parámetros no encontrado")
        
        with self.parameters_path.open('r', encoding="utf-8") as file:
            groups = [line.strip() for line in file if line.strip()]
            return groups
        
    
    @measure_run_time
    def save_groups(self, groups: list[str]) -> None:
        """
        The function `save_groups` writes a list of strings to a file specified by
        `self.parameters_path`.
        
        :param groups: The `groups` parameter is a list of strings containing the names of different
        groups that you want to save. The `save_groups` method takes this list of group names and writes
        them to a file specified by `self.parameters_path`. Each group name is written on a new line in
        the file
        :type groups: list[str]
        """
        with self.parameters_path.open('w', encoding='utf-8') as file:
            file.write('\n'.join(groups))

    
    @measure_run_time
    def add_group(self, entry_group: str) -> tuple[bool, str]:
        """
        The function `add_group` takes a group name as input, checks if it is valid and not already
        registered, then adds it to a list of groups and saves the updated list.
        
        :param entry_group: The `entry_group` parameter is a string representing the name of the group
        that is being added. It is expected to be stripped of any leading or trailing whitespace before
        processing
        :type entry_group: str
        :return: A tuple is being returned with a boolean value as the first element and a string
        message as the second element.
        """
        group_name = entry_group.strip()
        if not group_name:
            return False, "El nombre del grupo no puede estar vacío"

        groups = self.read_groups_from_file()
        if group_name in groups:
            return False, f"El grupo '{group_name}' ya está registrado"

        groups.append(group_name)
        self.save_groups(groups)

        return True, f"Grupo '{group_name}' registrado"
    

    @measure_run_time
    def remove_group(self, group_name: str) -> tuple[bool, str]:
        """
        This Python function removes a specified group from a list of groups and returns a tuple
        indicating success or failure along with a message.
        
        :param group_name: The `group_name` parameter is a string that represents the name of the group
        that you want to remove from the list of groups
        :type group_name: str
        :return: The function `remove_group` returns a tuple containing a boolean value and a string.
        The boolean value indicates whether the group was successfully removed (True if removed, False
        if the group does not exist), and the string provides a message indicating the outcome of the
        removal operation.
        """
        groups = self.read_groups_from_file()

        if group_name not in groups:
            return False, f"El grupo '{group_name}' no existe"
        
        groups.remove(group_name)
        self.save_groups(groups)

        return True, f"Grupo '{group_name}' eliminado"