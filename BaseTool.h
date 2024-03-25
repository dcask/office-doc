// BaseTool.h: interface for the CBaseTool class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_BASETOOL_H__9C41DDAD_BAFB_4426_AF30_06338C884120__INCLUDED_)
#define AFX_BASETOOL_H__9C41DDAD_BAFB_4426_AF30_06338C884120__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "ado2.h"

class CADOBaseTool
{
private:
	CADODatabase*	m_DataBase;		// объект базы
	CADORecordset*	m_RecordSet;	// объект ответа на запрос
	BOOL			m_bError;		// признак ошибки
	BOOL			m_bShowError;	// признак отображения сообщения с ошибкой
	CString			m_csErrorString;// текст последней ошибки
	CString			m_csDateFormat;	// формат даты для отображения методами класса
	CString			m_csProcNum;	// внутренний идентификатор метода
	DWORD			m_dwRecords;	// количество полученных записей
	long			m_lFields;		// количество полей в ответе
	long			m_lTime;		// время исполнения запроса
	int				m_iErrorPoint;	// точка в методе для  отладки
	static int		m_iObjCount;	// количество объектов данного класса
	CPtrArray		m_aRecordSet;

	CADOBaseTool();											// конструктор по умолчанию
public:
	void Dump(CString filename);
	int LoadComboBox(CWnd* cb, CADOCommand &pCmd);
	BOOL AddRecordSet();
	UINT GetRecordsCount(int rs_num=0);
	CADORecordset* GetRS(int n=0);
	BOOL IsError();
	
	CADOBaseTool(LPCTSTR connection,\
		LPCTSTR userId=_T(""), LPCTSTR password=_T(""));	// конструктор

	CADODatabase* GetDB() const {return m_DataBase;}		// указатель на объект базы
	CADORecordset* operator->();							// указатель на объект ответа
	
	void ThrowError(CADOException* e);						// обработчик ошибок
	void ThrowError(LPCTSTR error);							// обработчик ошибок
	void ShowError(BOOL show);								// показывать/скрывать ошибки
	void SetDateFormat(LPCTSTR format);						// установить формат даты
	
	BOOL Execute(LPCTSTR query,int rs_num=0);				// исполнить запрос
	BOOL Execute(CADOCommand *pCmd,int rs_num=0);			// исполнить запрос
	BOOL Compact();											// упаковать базу
	
	int LoadComboBox(CWnd* cb, LPCTSTR query);				// загрузить выпадающий список

	static UINT GetObjCount(){return m_iObjCount;}			// получить количество объектов
	
	CString GetStringValue(int index, int rs_num=0);		// получить значение поля
	CString GetStringValue(LPCTSTR fieldname, int rs_num=0);// получить значение поля
	
	long GetQueryTime() const {return m_lTime;}				// получить время исполнения
	
	virtual ~CADOBaseTool();								// деструктор

};
#endif // !defined(AFX_BASETOOL_H__9C41DDAD_BAFB_4426_AF30_06338C884120__INCLUDED_)
