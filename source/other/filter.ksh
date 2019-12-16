#!/bin/ksh

find .. -name 'modMain.bas' | xargs grep '\"*\"' > xxx.txt
#find . -name '*.frm' | xargs grep InitCheckBox >> xxx.txt
#find . -name '*.frm' | xargs grep lsvMaster.ListItems.Add >> xxx.txt
#find . -name '*.frm' | xargs grep trvMain.Nodes.Add >> xxx.txt