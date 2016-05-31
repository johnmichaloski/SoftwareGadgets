//
// Utils.cpp
//

#include "StdAfx.h"
#include "Utils.h"
#include <sys/types.h>

#include <errno.h>
#include <iostream>
#include <sstream>
#include "dirent.h"
#include "StdStringFcn.h"
#include <sys/stat.h>

static BOOL isFolder(LPCTSTR path)
{
    DWORD attr = GetFileAttributes(path);
    if (DWORD(-1) == attr) // path not exist
        return FALSE;
    return FILE_ATTRIBUTE_DIRECTORY & attr;
}

Utils::Utils(void)
{
}

Utils::~Utils(void)
{
}

std::vector<std::string> Utils::FileList(std::string folder)
{
	std::vector<std::string> files;

	DIR *dp;
	struct dirent *dirp;
	if((dp = opendir(folder.c_str())) == NULL) 
	{
		std::stringstream ss (std::stringstream::in);
		ss<< "Error(" << errno << ") opening " << folder << std::endl;
		throw ss.str();
	}

	if(folder[folder.size()-1]!='\\')
		folder+="\\";

	while ((dirp = readdir(dp)) != NULL) 
	{
		if(dirp->d_name[0]=='.')
			continue;
		std::string filename(dirp->d_name);
		filename=MakeLower(filename);
		if(filename.rfind("jpg") != std::string::npos)
			files.push_back(folder+std::string(dirp->d_name));
	}
	closedir(dp);
	return files;
}

std::vector<std::string> Utils::DirList(std::string folder)
{
	std::vector<std::string> folders;

	DIR *dp;
	struct dirent *dirp;
	if((dp = opendir(folder.c_str())) == NULL) 
	{
		std::stringstream ss (std::stringstream::in);
		ss<< "Error(" << errno << ") opening " << folder << std::endl;
		throw ss.str();
	}

	if(folder[folder.size()-1]!='\\')
		folder+="\\";

	std::string path (folder);
	while ((dirp = readdir(dp)) != NULL) 
	{
		if(dirp->d_name[0]=='.')
			continue;


		std::string foldername(dirp->d_name);
		foldername=MakeLower(foldername);

//		DWORD attr =  GetFileAttributes(foldername.c_str());
//
		//struct stat st;
		//stat(dirp->d_name, &st);
		//if(S_ISDIR(st.st_mode))
		if(isFolder((path+foldername).c_str()))
		{
			folders.push_back(path+foldername);
			std::vector<std::string> list = DirList(path+foldername);
			folders.insert( folders.begin(), list.begin(), list.end());
		}
	}
	closedir(dp);
	return folders;
}


void Utils::TrimExtension(std::vector<std::string>&strs, std::string extension)
{
	for(UINT i=0; i< strs.size(); i++)
	{
		strs[i]=ReplaceOnce(strs[i],extension, ""); 
	}
}
