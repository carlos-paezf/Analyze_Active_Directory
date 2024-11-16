from pyad import adgroup
from openpyxl import Workbook


def get_group_members(group_name) -> list:
    """
    The function `get_group_members` retrieves the members of a specified Active Directory group by its
    name.
    
    :param group_name: The `get_group_members` function takes a `group_name` parameter as input, which
    is expected to be a string representing the name of a group. The function then attempts to retrieve
    the members of the specified group using Active Directory operations. If successful, it returns a
    list of member common names (
    :return: The function `get_group_members(group_name)` is returning a list of common names (cn) of
    the members of the Active Directory group specified by `group_name`. If an error occurs during the
    process, an empty list is returned.
    """
    try:
        group = adgroup.ADGroup.from_cn(group_name)
        members = group.get_members()
        return [member.cn for member in members]
    except Exception as e:
        print(f"Se ha encontrado un error: {e}")
        return []
    

def save_to_excel(group_members: dict, output_file: str) -> None:
    """
    The function `save_to_excel` takes a dictionary of group members and saves the data to an Excel file
    with each group as a column and members as rows.
    
    :param group_members: The `group_members` parameter is a dictionary where the keys represent group
    names and the values are lists of group members belonging to each group
    :type group_members: dict
    :param output_file: The `output_file` parameter in the `save_to_excel` function is a string that
    represents the file path where the Excel file will be saved. This parameter should include the file
    name and extension (e.g., "output.xlsx") to specify the location where the Excel file will be
    created or updated
    :type output_file: str
    """
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Grupos y Miembros"

    headers = list(group_members.keys())
    
    max_members = max(len(members) for members in group_members.values())

    for col, group_name in enumerate(headers, start=1):
        sheet.cell(row=1, column=col, value=group_name)

        for row in range(max_members):
            for col, group_name in enumerate(headers, start=1):
                members = sorted(group_members[group_name])

                if row < len(members):
                    sheet.cell(row=row + 2, column=col, value=members[row])

    workbook.save(output_file)
    print(f"Datos guardados en {output_file}")


if __name__ == "__main__":
    group_list = []

    group_members = {}

    for group_name in group_list:
        print(f"Consultado miembros del grupo {group_name}")
        members = get_group_members(group_name)
        group_members[group_name] = members

    output_file = "Grupos_y_Miembros_AD.xlsx"
    save_to_excel(group_members, output_file)