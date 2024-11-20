from pyad import adgroup
from .group_manager import Group_Manager

from measure_run_time import measure_run_time


class Search_Active_Directory():
    def __init__(self, group_list: list[str]):
        self.group_members = self.check_active_directory(group_list)

    
    def get_group_members(self, group_name: str) -> list[str]:
        """
        This function retrieves the members of a specified Active Directory group by its name.
        
        :param group_name: The `get_group_members` function takes a `group_name` parameter, which is a
        string representing the name of a group. This function attempts to retrieve the members of the
        specified group using the `adgroup.ADGroup.from_cn(group_name)` method and then returns a list
        of member common names (`
        :type group_name: str
        :return: A list of strings containing the common names (cn) of the members of the specified
        group. If an error occurs during the process, an empty list will be returned.
        """
        try:
            group = adgroup.ADGroup.from_cn(group_name)
            members = group.get_members()
            return [member.cn for member in members]
        except Exception as e:
            print(f"Se ha encontrado un error: {e}")
            return []
        
    
    @measure_run_time
    def check_active_directory(self, group_list: list[str]) -> dict[str, list[str]]:
        """
        The function `check_active_directory` takes a list of group names, retrieves the members of each
        group, and returns a dictionary mapping group names to their respective members.
        
        :param group_list: The `check_active_directory` method takes a list of group names as input and
        returns a dictionary where the keys are group names and the values are lists of group members
        :type group_list: list[str]
        :return: A dictionary is being returned, where the keys are group names (strings) and the values
        are lists of group members (strings).
        """
        group_members = {}

        for group_name in group_list:
            print(f"Consultando miembros del grupo '{group_name}'")
            members = self.get_group_members(group_name)
            if len(members) == 0:
                print(f"\tNo se han encontrado miembros para el grupo '{group_name}'")
            else:
                group_members[group_name] = members
        
        return group_members
 
