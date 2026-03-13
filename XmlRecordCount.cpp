#include "stdafx.h"
#include <fcntl.h>
#include <stdio.h>
#include <io.h>
#include <sys/stat.h>

#define BUFSIZE 8192
#define ITERATION_COUNT 20 

BOOL isApprox(long cbSeek)
{
	if(cbSeek < BUFSIZE * 2)
		return FALSE;
	else
		return TRUE;
}

HRESULT XMLRecordCountLong(BSTR xmlFilename, ULONG parlevel, ULONG *nRecCount)
{
	int handle;
	int i=0;
	int cbReadBytes = 0;

	long  cbFileSize;
	long  cbHeader = 0;
	long  cbTotalLen = 0;
	long  cbSeek = 0;

	char *buffer = NULL;
	char  chPreChar = 0;
	char  chPrePreChar = 0;
	char *achar;

	unsigned long rcount = 0;
	int level = 0;
	bool comment = false;

	char *szTag = NULL;
	int   cbTag = 0;
    
	*nRecCount = 0;

	buffer = new char[BUFSIZE];
	memset(buffer, 0, BUFSIZE);
	
	errno_t erc = _wsopen_s(&handle, xmlFilename, _O_RDONLY | _O_BINARY, _SH_DENYWR, _S_IREAD);  
	if(erc)
		return STG_E_FILENOTFOUND;

	cbFileSize = _lseek(handle, 0L, SEEK_END);
	cbSeek = (cbFileSize - BUFSIZE - (BUFSIZE * ITERATION_COUNT)) / (ITERATION_COUNT - 1);
	_lseek(handle, 0L, SEEK_SET);

	do
	{
		cbReadBytes = _read(handle, buffer, BUFSIZE);
		for(i=0, achar = buffer; i < cbReadBytes; i++)
		{
			if(*achar == '?')		// объявление xml
			{
				if(chPreChar == '<')
					comment = true;
			}
			else if(*achar == '/')		// </
			{
				if(chPreChar == '<' && (!comment))
				{
					if(level == parlevel)
					{
						rcount++;		
					}
					level--;
				}
			}
			else if(*achar == '>')
			{
				if(chPreChar == '/' && (!comment))					// />
				{
					if(level == parlevel)
					{
						rcount++;
					}
					level--;
				}
				else if(chPreChar == '-' && chPrePreChar == '-')		// -->
				{
					comment = false;
				}
				else if(chPreChar == '?')
				{
					comment = false;
				}
				else
				{
					if(level == parlevel)
					{
						if(rcount == 0 && isApprox(cbSeek))
						{
							szTag = new char[cbTotalLen + 2];
							cbTag = cbTotalLen + 1;
							memmove(szTag, achar - cbTotalLen, cbTag);
							szTag[cbTag] = 0;
							cbReadBytes = 0;
							break;
						}
					}
				}
			}
			else if(*achar == '-')
			{
				if(chPreChar == '!' && chPrePreChar == '<')		// <!-
					comment = true;
			}
			else if(*achar != '!' && *achar != '?' && chPreChar == '<')			// <*
			{
				if(!comment)
				{
					level++;
					if(level == parlevel)
					{
						if(cbHeader==0)
							cbHeader = cbTotalLen - 1;
						cbTotalLen = 1;
					}
				}
			}
			chPrePreChar = chPreChar;
			chPreChar = *achar;
			cbTotalLen++;
			achar++;
		}
	}
	while(cbReadBytes > 0);
	
	if(!isApprox(cbSeek))
	{
		*nRecCount = rcount;
	}
	else
	{
		cbSeek = (cbFileSize - cbHeader - (BUFSIZE * ITERATION_COUNT)) / (ITERATION_COUNT - 1);

		cbTotalLen = 0;
		level = 0;
		rcount = 0;
		_lseek(handle, cbHeader, SEEK_SET);
		do
		{
			memset(buffer, '\0', BUFSIZE);
			cbReadBytes = _read(handle, buffer, BUFSIZE);
			cbReadBytes -= cbTag;
			char *pchPreChar = NULL;
			int cbElementsLen = 0;
			rcount = 0;
			for(i=0, achar=buffer; i < cbReadBytes; i++, achar++)
			{
				if(strncmp(achar, szTag, cbTag)==0)
				{
					if(pchPreChar)
					{
						cbElementsLen += (achar - pchPreChar);
						rcount++;
						//printf("%d %d\n", (achar - pchPreChar), rcount);
					}
					pchPreChar = achar;
				}
			}
			if(rcount > 0)
			{
				cbTotalLen += (cbElementsLen / rcount);
				level++;
			}
			_lseek(handle, cbSeek, SEEK_CUR);
		}
		while(cbReadBytes > 0);
		if(level > 0) 
			*nRecCount = (cbFileSize - cbHeader) / (cbTotalLen / level);
	}
	_close(handle);

	if(szTag)
		delete [] szTag;
	delete buffer;
	
	//printf("rcount=%u\n", *nRecCount);
	return S_OK;
}

HRESULT XMLRecordCount(BSTR xmlFilename, ULONG parlevel, ULONG *nRecCount)
{
	int   i=0;
	char *buffer = NULL;
	char  preChar = 0;
	char  prePreChar = 0;
	char *achar;
	DWORD n = 0;	
	unsigned long rcount = 0;
	int level = 0;
	bool comment = false;
    
	*nRecCount = 0;

	buffer = new char[BUFSIZE];
	memset(buffer, 0, BUFSIZE);
	
	WideCharToMultiByte(CP_ACP, 0, xmlFilename,-1, buffer, BUFSIZE, 0, 0);
	OFSTRUCT ofstruct = { sizeof(OFSTRUCT) };
	HFILE hFile = OpenFile(buffer, &ofstruct, OF_READ);

	if(hFile == HFILE_ERROR)
		return STG_E_FILENOTFOUND;

	memset(buffer, '\0', BUFSIZE);
	n = 1;
	DWORD lf = 0;

	while(n > 0)
	{
		memset(buffer, '\0', BUFSIZE);
		if(!ReadFile((HANDLE)hFile, buffer, BUFSIZE, &n, NULL))
			break;
		lf+=n;
		for(i=0, achar = buffer; i < n; i++)
		{
			if(*achar == '?')		// объявление xml
			{
				if(preChar == '<')
					comment = true;
			}
			else if(*achar == '/')		// </
			{
				if(preChar == '<' && (!comment))
				{
					if(level == parlevel)
						rcount++;
					level--;
				}
			}
			else if(*achar == '>')
			{
				if(preChar == '/' && (!comment))					// />
				{
					if(level == parlevel)
						rcount++;
					level--;
				}
				else if(preChar == '-' && prePreChar == '-')	// -->
				{
					comment = false;
				}
				else if(preChar == '?')
				{
					comment = false;
				}
			}
			else if(*achar == '-')
			{
				if(preChar == '!' && prePreChar == '<')		// <!-
					comment = true;
			}
			else if(*achar != '!' && *achar != '?' && preChar == '<')			// <*
			{
				if(!comment)
				{
					level++;
				}
			}
			prePreChar = preChar;
			preChar = *achar;
			achar++;
		}
	}
	::CloseHandle((HANDLE)hFile);
	delete buffer;
	
	*nRecCount = rcount;
	return S_OK;
}

