#pragma once
template <class T> class TList
{
protected:
	class TELEMENT
	{
	public:
		TELEMENT(T * pT)
		{
			ptr = pT;
			next = NULL;
		}
		TELEMENT(T * pT, int lelem)
		{
			BYTE *tempPtr = new BYTE[lelem];
			memmove(tempPtr, pT, lelem);
			ptr = (T*)tempPtr;
			next = NULL;
		}
		~TELEMENT()
		{
			delete ptr;
		}
		T * ptr;
		TELEMENT * next;
	} *m_ptr;
protected:
	int m_nCount;
	TELEMENT * m_pCurrentElement;
public:
	TList()
	{
		m_ptr = NULL;
		m_nCount = 0;
		m_pCurrentElement = NULL;
	}
		
	~TList()
	{
		TELEMENT * cpt = m_ptr;
		TELEMENT * last;
		for (int i = 0; i < m_nCount; i++)
		{
			last = cpt->next;
			delete cpt;
			cpt = last;
		}
		m_nCount = 0;
	}
	int Add(T *pT)
	{
		if(m_ptr == NULL)
		{
			m_ptr = new TELEMENT(pT);
			m_nCount = 1;
			return 1;
		}
		else
		{
			TELEMENT * cpt = m_ptr;
			for(int i = 0; i < m_nCount - 1; i++)
			{
				cpt = cpt->next;
			}
			cpt->next = new TELEMENT(pT);
			m_nCount++;
			return m_nCount;
		}
	}
	int Add(T * pT, int lel)
	{
		if(m_ptr == NULL)
		{
			m_ptr = new TELEMENT(pT, lel);
			m_nCount = 1;
			return 1;
		}
		else
		{
			TELEMENT * cpt = m_ptr;
			for(int i = 0; i < m_nCount - 1; i++)
			{
				cpt = cpt->next;
			}
			cpt->next = new TELEMENT(pT, lel);
			m_nCount++;
			return m_nCount;
		}
	}

	T* Next()
	{
		if(m_pCurrentElement == NULL)
		{
			m_pCurrentElement = m_ptr;
			if(!m_pCurrentElement)
				return NULL;
			return m_pCurrentElement->ptr;
		}
		else
		{
			TELEMENT * cptr = m_pCurrentElement->next;
			if(cptr == NULL)
				return NULL;
			m_pCurrentElement = cptr;
			return cptr->ptr;
		}
	}
	virtual BOOL Compare(T *ptr, void *key)
	{
		if(ptr == (T*)key)
			return TRUE;
		else
			return FALSE;
	}
	T* Find(void *ptr)
	{
		TELEMENT * cpt = m_ptr;
		for(int i = 0; i < m_nCount; i++)
		{
			if(Compare(cpt->ptr, ptr) == TRUE)
			{
				return cpt->ptr;
			}
			cpt = cpt->next;
		}
		return NULL;
	}
	void Remove(T *ptr)
	{
		if(!m_ptr)
			return;
		TELEMENT *cpt = m_ptr;
		TELEMENT *last = NULL;
		for(int i = 0; i < m_nCount; i++)
		{
			if(cpt->ptr == ptr)
			{
				if(last)
					last->next = cpt->next;
				if(cpt == m_ptr)
					m_ptr = cpt->next;
				delete cpt;
				if(last)
					m_pCurrentElement = last;
				else
					m_pCurrentElement = NULL;
				m_nCount--;
				break;
			}
			last = cpt;
			cpt = cpt->next;
		}
	}

	void MoveFirst()
	{
		m_pCurrentElement = NULL;
	}
	int Count()
	{
		return m_nCount;
	}
};