#!/usr/bin/env python

"""
This script trims the size of Microsoft Office Mac 2016 
installation. On my system, I am reduce the size of the
installation by ~60%

*** Run this script with sudo privileges ***

This code is provided as is. I am not responsible for any
loss of files or damage to your system, or your installation
of Office. 

--Deepak

------------------------------------------------------------

It's disappointing to see that the latest Office for Mac
is a very bloated distribution. A lot of the bloat is due
to language files and duplicates for font directory and
profiling tools between the apps. 

Size after fresh install:
Word     1.6 GB
Excel    1.5 GB
PPT      1.4 GB
Outlook  1   GB
One Note 692 MB
---------------
Total    6.2 GB

After running Monolingual to remove language files:
Word     1.3 GB
Excel    1.3 GB
PPT      1.2 GB
Outlook  848 GB
One Note 570 MB

After this script
Word     953 MB
Excel    451 MB
PPT      432 MB
Outlook  310 MB
One Note 264 MB
---------------
Total    2.4 GB

"""

import os.path as osp
import os, shutil



WD = "/Applications/Microsoft Word.app"
XL = "/Applications/Microsoft Excel.app"
PP = "/Applications/Microsoft PowerPoint.app"
OT = "/Applications/Microsoft Outlook.app"
ON = "/Applications/Microsoft OneNote.app"

apps = [WD, XL, PP, OT, ON]
except_WD = [XL, PP, OT, ON]


def get_size(start_path = '.'):
    # http://stackoverflow.com/questions/1392413/calculating-a-directory-size-using-python
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(start_path):
        for f in filenames:
            fp = os.path.join(dirpath, f)
            total_size += os.path.getsize(fp)
    return total_size

before = 0.0
for app in apps:
    a = get_size(app)/1024./1024./1024.
    before += a
    print app, a

print "Total before: ", before

for app in except_WD:
    print(app)
    os.chdir(osp.join(app, "Contents/Resources/Fonts"))
    fonts = os.listdir(".")
    for font in fonts:
        print(font)
        word_file = osp.join(WD,"Contents/Resources/Fonts", font)
        if osp.exists(word_file):
            os.remove(font)
            os.symlink(word_file, font)


after = 0.0
for app in apps:
    a = get_size(app)/1024./1024./1024.
    after += a
    print app, a


# Removing profiling tools
print("Removing profiling tools")
for app in apps:
    os.chdir(osp.join(app, "Contents/SharedSupport/Proofing Tools"))
    langs = os.listdir(".")
    for lang in langs:
        if not "English" in lang:
            shutil.rmtree(lang)
        

print "Total after: ", after
