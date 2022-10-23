#!/usr/bin/python2
#coding=utf-8
import os



# no include ".pdf"
newnamelen = 17

path='.'
filelist=os.listdir(path)
#print(filelist)

g_newnamelist=[]

def rename(oldname, filetypestr):

	if (fname.endswith(filetypestr)>0) and (len(oldname) > newnamelen+len(filetypestr)):
		newname = fname[0:newnamelen]+filetypestr
		if (newname in g_newnamelist) or (newname in filelist):
			print("[!] ["+ newname + "] already exist.   ["+ oldname + "] cannot be renamed. ")
			return
		os.rename(oldname, newname)
		print("mv "+oldname+" "+ newname)
		g_newnamelist.append(newname)

for fname in filelist:
	#print("scan : "+fname)
	#print(type(fname))
	rename(fname,".pdf")
	rename(fname,".xls")
	rename(fname,".xlsx")




