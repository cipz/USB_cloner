# Cloner for inserted USB and other media cards

Script to automatically save the contents of USB and SD cards upon insertion in the computer.
For each USB driver it creates one folder named "ComputerName_USBVolumeName_SerialNumber" in the `%AppData%` folder.

Here there are two versions of the sript:
- `copy_all_contents.vbs` which copies all the content of the USB / media drive
- `copy_pdf_documents.vbs` which copies all `pdf` / `doc` / `docx` documents in the USB / media drive

This repo is based on the answer to [this old stackoverflow question](https://stackoverflow.com/questions/34979009/hidden-script-to-copy-content-of-usb).
Kudos to [Hackoo](https://stackoverflow.com/users/3080770/hackoo) for answering my question with this awesome script.

Examples on how to use it can be:
- backing up files upon usb insertion (and start other automatization)
- download pictures from the SD of a camera without copying them manually after insertion
- ~~let it run in the background as a daemon and copy the files that people have on their USB drive~~
- etc.
