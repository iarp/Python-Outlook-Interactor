import win32com.client


class FolderNotFoundException(Exception):
    pass


class TooManyFoldersFoundException(Exception):
    pass


class Outlook:

    namespace = None
    application = None

    folder_structure = {}

    config = {}

    def __init__(self):
        self.application = win32com.client.Dispatch('Outlook.Application')
        self.namespace = self.application.GetNamespace('MAPI')

        self.folder_structure = self.get_folder_structure()

    def create_email(self, subject, body, to, cc=None, attachments=None, display_on_creation=True,
                     send_immediately=False):

        message = self.application.CreateItem(win32com.client.constants.olMailItem)
        message.To = to
        message.Subject = subject

        if cc:
            message.CC = cc

        # Must display the email first for the signature to populate, then we can write in it.
        message.Display()

        message.HTMLBody = body + message.HTMLBody

        if isinstance(attachments, list):
            for attachment in attachments:
                message.Attachments.Add(attachment)

        message.Save()

        if not display_on_creation:
            message.Close(True)
        if send_immediately:
            message.Send()

    def get_folder_structure(self, folder=None, data=None):
        """ Recursive function that loads all folders found in Outlook on the running computer.

        {
            'Outlook Data File': {
                'id': '12412h1h4u124jh......',
                'folders': {
                    'Inbox': {
                        'id': '...............',
                        'folders': recursive directory tree.
                    }
                }
            }
        }

        :param folder: Folder object where to start structure building from. Default is namespace (root)
        :param data: Default data for the dict being built?
        :return:
        """
        if not folder:
            folder = self.namespace
        if not isinstance(data, dict):
            data = {}

        for f in folder.Folders:

            folder_name = f.Name

            if folder_name not in data:
                data[folder_name] = {'id': f.EntryID, 'folders': []}

            if f.Folders:
                data[folder_name]['folders'] = self.get_folder_structure(f)

        return data

    def get_folder_by_id(self, folder_id):
        return self.namespace.GetFolderFromID(folder_id)

    def get_folder_by_tree(self, wanted_folder_structure: [tuple, list]):
        """ Returns Folder object based on tuple or list of folder names
                given single-tree style.

        Using the dict example found in self.get_folder_structure method,
            to obtain the Inbox you would do the following

            o.get_folder_by_tree(['Outlook Data File', 'Inbox'])

        """

        # Since it's a recursive dict, start off at the top
        data = o.folder_structure
        key = None

        # If there only one primary data file in Outlook, doing this
        # allows us to skip the need of requiring the data files name
        # too, this way you can just pass the folder structure alone.
        if len(data) == 1:

            key = next(iter(data))
            data = o.folder_structure[key]['folders']

        # Go over each folder wanted, in order, overwriting the data
        # dict used with the next level down each time
        for index, folder_name in enumerate(wanted_folder_structure):

            # If our key matches the folder name, skip this iteration
            # this only happens at the data file name level anyways.
            if not index and key == folder_name:
                continue

            # Ensure the folder we're after is in the current data dict
            if folder_name in data:

                # Overwrite the data dict with the next levels information
                data = data[folder_name]

                # If we're NOT at the last iteration in the loop then
                # we want to pass the next levels information folder structure only.
                if index + 1 != len(wanted_folder_structure):
                    data = data['folders']
            else:
                raise FileNotFoundError('Folder "{}" was not found in the tree given! {}'.format(folder_name,
                                                                                                 wanted_folder_structure))

        return self.get_folder_by_id(data['id'])

    def find_folder_by_name(self, folder_name: str):
        """ Finds a folder in outlook by a single name alone.

        First it recursively searches the folder structure for all folders with the matching name

        If 0 folders match folder_name value a FolderNotFoundException is raised
        If more than 1 folder is found with matching folder_name value a TooManyFoldersFoundException is raised
        Otherwise we return the folder object.

        """

        # We need to find out how many times the folders name is found in the open data files.
        found = self._loop_folder_finder(folder_name, o.folder_structure)

        if not found:
            raise FolderNotFoundException('Folder {} was not found in Outlook'.format(folder_name))
        elif found > 1:
            raise TooManyFoldersFoundException('Folder {} was found more than once, use get_folder_by_tree with tree.'.format(folder_name))

        folder_id = self._loop_folder_finder(folder_name, o.folder_structure, return_id=True)

        return self.get_folder_by_id(folder_id=folder_id)

    def _loop_folder_finder(self, folder_looking_for: str, structure: dict, counter=0, return_id=False):
        """ Recursively attempts to find folder_looking_for value in the structure dict

        If return_id=True then the folders ID is returned, otherwise a counter of how
            many times that folder was in Outlook is returned.

        """
        if not structure:
            return 0

        for k, d in structure.items():

            if k == folder_looking_for:
                counter += 1

                if return_id:
                    return d['id']

            # If there are more folders to search through, keep going deeper.
            if d['folders']:

                returned_data = self._loop_folder_finder(
                    folder_looking_for=folder_looking_for,
                    structure=d['folders'],
                    return_id=return_id)

                # If the value returned is not an integer, return it immediately because it's the ID
                if not isinstance(returned_data, int):
                    return returned_data

                counter += returned_data

        return counter

if __name__ == '__main__':
    from pprint import pprint
    import sys

    o = Outlook()

    folder_object = o.get_folder_by_tree(('Outlook Data File', 'P1', 'C1', 'GC1'))

    print(folder_object.Name)
    for x in range(folder_object.Items.Count, 0, -1):
        email = folder_object.Items.Item(x)

        print(email.Subject)
        print(email.Body)
        print(email.HTMLBody)
