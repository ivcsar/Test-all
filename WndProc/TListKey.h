#include "..\TList.h"

template <class T, class _tKey> class TListKey : public TList <T>
{
public:
	void RemoveByKey(_tKey key)
	{
		if(!m_ptr)
			return;
		
		TELEMENT * cpt = m_ptr;
		TELEMENT * last = NULL;
		for(int i = 0; i < m_nCount; i++)
		{
			if(cpt->ptr->Compare(key) == 0)
			{
				if(last)
					last->next = cpt->next;
				delete cpt;
				m_nCount--;
				break;
			}
			last = cpt;
			cpt = cpt->next;
		}
		if(m_nCount == 0)
		{
			m_pCurrentElement = NULL;
			m_ptr = NULL;
		}
	}
	T *FindByKey(_tKey key)
	{
		TELEMENT * cpt = m_ptr;
		for(int i = 0; i < m_nCount; i++)
		{
			if(cpt->ptr->Compare(key) == 0)
			{
				return cpt->ptr;
			}
			cpt = cpt->next;
		}
		return NULL;
	}
};