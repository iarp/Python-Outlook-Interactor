Python Outlook Interactor
=========================

This is a class I wrote to assist my programs when interacting with Outlook.

I was using Outlook 2013/2016 but have tried running this on 2010 without errors so far.


    o = Outlook()

    folder_object = o.get_folder_by_tree(('Outlook Data File', 'Forms'))

    # See the following link for available properties and methods.
    # https://msdn.microsoft.com/en-us/library/office/ff861332.aspx

    for x in range(folder_object.Items.Count, 0, -1):
        email = folder_object.Items.Item(x)

        print(email.Subject)
        print(email.Body)
        print(email.HTMLBody)
