// ADO.cpp: implementation of the ADO class.
// 
//////////////////////////////////////////////////////////////////////
// Copyright (c) 2004 by cdecl (byung-kyu kim)
// EMail : cdecl@interpark.com
//////////////////////////////////////////////////////////////////////
//#include "StdAfx.h"			// MFC Project
#include "ADO.h"
using namespace ADODB;

using namespace GLASS;
//////////////////////////////////////////////////////////////////////
// ADO Implemetation
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// Construction
ADO::ADO()
{
	pConnection_ = NULL;
	pRecordset_ = NULL;
}

//////////////////////////////////////////////////////////////////////
// Destructor
ADO::~ADO()
{
	Release();
}

//////////////////////////////////////////////////////////////////////
// Member function

//////////////////////////////////////////////////////////////////////
// Free Memory
void ADO::Release()
{
	if (pRecordset_) {
		if (pRecordset_->State == adStateOpen) {
			pRecordset_->Close();
		}
		pRecordset_.Release();
	}

	if (pConnection_) {
		if (pConnection_->State == adStateOpen) {
			pConnection_->Close();
		}		
		pConnection_.Release();
	}
}


//////////////////////////////////////////////////////////////////////
// ODBC Connection and Recordset CreateInstance
void ADO::Create(const _bstr_t &bstrConnectionString, const _bstr_t &bstrUID, const _bstr_t &bstrPWD)
{
	pConnection_.CreateInstance(__uuidof(Connection));
	pConnection_->Open(bstrConnectionString, bstrUID, bstrPWD, -1L);		
}


//////////////////////////////////////////////////////////////////////
// Recordset Open
void ADO::OpenRs(const _bstr_t &bstrSql, CursorLocationEnum adCursorLocation, const long lCommandTimeout)
{
	if (pRecordset_ == NULL) {
		pRecordset_.CreateInstance(__uuidof(Recordset));
	}
	
	// CommandTimeout  
	if (lCommandTimeout >= 0) {
		pConnection_->PutCommandTimeout(lCommandTimeout);
	}

	_variant_t varConnection;
	varConnection = (IUnknown *)pConnection_;

	pRecordset_->CursorLocation = adCursorLocation;
	pRecordset_->Open (
			bstrSql,
			varConnection,
			adOpenForwardOnly,
			adLockUnspecified,
			-1L
	);

	if (adCursorLocation != adUseServer) {
		pRecordset_->PutRefActiveConnection(NULL); 
	}
}

//////////////////////////////////////////////////////////////////////
// Query Excute
long ADO::Execute(const _bstr_t &bstrSql, const long lCommandTimeout)
{
	_variant_t varRecordsAffected;

	// CommandTimeout  
	if (lCommandTimeout >= 0) {
		pConnection_->PutCommandTimeout(lCommandTimeout);
	}

	pConnection_->Execute (
				bstrSql,
				&varRecordsAffected,
				-1L
	);

	return (long)varRecordsAffected;
}




//////////////////////////////////////////////////////////////////////
// Construction
ADOComm::ADOComm() : ADO()
{
}

//////////////////////////////////////////////////////////////////////
// Destructor
ADOComm::~ADOComm() 
{
}


//////////////////////////////////////////////////////////////////////
// Recordset Open
void ADOComm::OpenRs(const _bstr_t &bstrSql, CursorLocationEnum adCursorLocation, const long lCommandTimeout, bool bPrepared)
{
	ADODB::_CommandPtr pCommand;
	pCommand.CreateInstance(__uuidof(Command));

	pCommand->PutCommandType(ADODB::adCmdText);
	pCommand->PutCommandText(bstrSql);
	pCommand->PutPrepared(bPrepared ? VARIANT_TRUE : VARIANT_FALSE);
	
	if (lCommandTimeout >= 0) {
		pCommand->PutCommandTimeout(lCommandTimeout);
	}

	pCommand->PutRefActiveConnection(pConnection_);
	pConnection_->PutCursorLocation(adCursorLocation);

	_variant_t varRecordsAffected;
	pRecordset_ = pCommand->Execute(&varRecordsAffected, &vtMissing, ADODB::adCmdText);
	pCommand->PutRefActiveConnection(NULL);

	if (adCursorLocation != adUseServer) {
		pRecordset_->PutRefActiveConnection(NULL); 
	}
}


void ADOComm::OpenRs(ADODB::_CommandPtr &pCommand, CursorLocationEnum adCursorLocation, const long lCommandTimeout)
{
	if (lCommandTimeout >= 0) {
		pCommand->PutCommandTimeout(lCommandTimeout);
	}

	pCommand->PutRefActiveConnection(pConnection_);
	pConnection_->PutCursorLocation(adCursorLocation);

	_variant_t varRecordsAffected;
	pRecordset_ = pCommand->Execute(&varRecordsAffected, &vtMissing, pCommand->GetCommandType());
	pCommand->PutRefActiveConnection(NULL);

	if (adCursorLocation != adUseServer) {
		pRecordset_->PutRefActiveConnection(NULL); 
	}
}




//////////////////////////////////////////////////////////////////////
// Query Excute
long ADOComm::Execute(const _bstr_t &bstrSql, const long lCommandTimeout, bool bPrepared)
{
	ADODB::_CommandPtr pCommand;
	pCommand.CreateInstance(__uuidof(Command));

	// CommandTimeout  
	if (lCommandTimeout >= 0) {
		pCommand->PutCommandTimeout(lCommandTimeout);
	}

	pCommand->PutCommandType(ADODB::adCmdText);
	pCommand->PutCommandText(bstrSql);
	pCommand->PutPrepared(bPrepared ? VARIANT_TRUE : VARIANT_FALSE);
	
	_variant_t varRecordsAffected;
	pCommand->PutRefActiveConnection(pConnection_);
	pCommand->Execute(&varRecordsAffected, &vtMissing, ADODB::adExecuteNoRecords); // NoRecords
	pCommand->PutRefActiveConnection(NULL);

	return (long)varRecordsAffected;
}


long ADOComm::Execute(ADODB::_CommandPtr &pCommand, const long lCommandTimeout)
{
	// CommandTimeout  
	if (lCommandTimeout >= 0) {
		pCommand->PutCommandTimeout(lCommandTimeout);
	}

	_variant_t varRecordsAffected;
	pCommand->PutRefActiveConnection(pConnection_);
	pCommand->Execute(&varRecordsAffected, &vtMissing, ADODB::adExecuteNoRecords); // NoRecords	
	pCommand->PutRefActiveConnection(NULL);   

	return (long)varRecordsAffected;
}




//////////////////////////////////////////////////////////////////////
// Copyright (c) 2004 by cdecl (byung-kyu kim)
// EMail : cdecl@interpark.com
//////////////////////////////////////////////////////////////////////
