# Office16CustomInstaller
Install Microsoft Office 2016 individual applications.
# Preconditions
* To use the installer, python3.x must be installed.
* Microsoft Office 2016 disc image(.iso).
# How to use
* Mount/extract iso file.
* Copy setup.py into directory that iso file mounted/extracted.
* run `python3 setup.py -h` to see usage.

```
python3 setup.py -h
usage: setup.py [-h] [-k ACTION] [-p PRODUCT] [-e EDITION] [-l LANG]

Microsoft Office 2016 downloader/installer example: python setup.py --action
install --product word --edition 64 --lang zh-cn

optional arguments:
  -h, --help  show this help message and exit
  -k ACTION   install | download
  -p PRODUCT  product to install, e.g. Word/PowerPoint/Excel/Outlook/Access/In
              foPath/Groove/Lync/OneDrive/Published/Project/Visio/SharePointDe
              signer
  -e EDITION  product edition, e.g. 64/32
  -l LANG     install language, e.g. en-us/zh-cn
```
# Basic steps
* Create a directory
`mkdir office2016`
* Change to the created directory
`cd office2016`
* run install command
`python3 setup.py -k install -l zh-cn -p Word -e 64`
Only install Word with Simple Chinese language.
