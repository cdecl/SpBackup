// ADO.h: interface for the ADO class.
//
//////////////////////////////////////////////////////////////////////
// Copyright (c) 2004 by cdecl (byung-kyu kim)
// EMail : cdecl@interpark.com
//////////////////////////////////////////////////////////////////////


#ifndef		_ADO_H_CDECL_
#define		_ADO_H_CDECL_

#import		"C:\Program Files\Common Files\System\ado\msado15.dll"	\
								rename("EOF", "adoEOF")

#include	<tchar.h>
#include	<comdef.h>				// using _bstr_t class 


namespace GLASS {
	

////////////////////////////////////////////////////////
// ADO class 
class ADO  
{
public:
	ADO();
	virtual ~ADO();

public:
	virtual long GetRecordCount() { return pRecordset_->GetRecordCount(); }
	virtual void MovePrevious() { pRecordset_->MovePrevious(); }
	virtual void MoveLast() { pRecordset_->MoveLast(); }
	virtual void MoveFirst() { pRecordset_->MoveFirst(); }
	virtual void MoveNext() { pRecordset_->MoveNext(); }

	virtual bool NextRecordset() { 
		_variant_t varRecordAffected;
		pRecordset_ = pRecordset_->NextRecordset(&varRecordAffected);

		if (pRecordset_ == NULL) {
			return false;
		}
		return true;
	}

	
	virtual bool IsEOF() const { 
		return (pRecordset_->adoEOF == VARIANT_TRUE); 
	};

	virtual bool IsBOF() const  { 
		return (pRecordset_->BOF == VARIANT_TRUE); 
	}

	virtual void Create(const _bstr_t &bstrConnectionString, const _bstr_t &bstrUID = _T(""), const _bstr_t &bstrPWD = _T(""));

	virtual void OpenRs(const _bstr_t &bstrSql, ADODB::CursorLocationEnum adCursorLocation = ADODB::adUseClient, const long lCommandTimeout = -1);

	virtual void CloseRs() {
		if (pRecordset_->State == ADODB::adStateOpen)
			pRecordset_->Close();
	}

	virtual long Execute(const _bstr_t &bstrSql, const long lCommandTimeout = -1);
	virtual void Release();

	virtual _variant_t Detach() { 
		return pRecordset_.Detach(); 
	}

	virtual _variant_t GetItem(const _bstr_t &bstrItem) const {
		return pRecordset_->GetCollect(bstrItem).vt == VT_NULL 
					? "" : pRecordset_->GetCollect(bstrItem);
	}

	virtual _variant_t GetItem(const long nItemIndex) const {
		return pRecordset_->GetCollect(nItemIndex).vt == VT_NULL 
					? "" : pRecordset_->GetCollect(nItemIndex);
	}

	virtual _variant_t operator()(const _bstr_t &bstrItem) const {	
		return GetItem(bstrItem); 
	}

	virtual _variant_t operator()(const long nItemIndex) const {	
		return GetItem(nItemIndex); 
	}

	virtual long BeginTran() {
		return pConnection_->BeginTrans();
	}

	virtual void CommitTran() {
		pConnection_->CommitTrans();
	}

	virtual void RollbackTran() {
		pConnection_->RollbackTrans();
	}

private:
	ADO& operator=(const ADO &ado);
	ADO(const ADO &ado);


public:
	ADODB::_RecordsetPtr pRecordset_;

protected:
	ADODB::_ConnectionPtr pConnection_;

};


class CommandHelper;

class ADOComm : public ADO
{
public:
	ADOComm();
	virtual ~ADOComm();

public:
	virtual void OpenRs(const _bstr_t &bstrSql, ADODB::CursorLocationEnum adCursorLocation = ADODB::adUseClient, const long lCommandTimeout = -1, bool bPrepared = false);
	virtual void OpenRs(ADODB::_CommandPtr &pCommand, ADODB::CursorLocationEnum adCursorLocation = ADODB::adUseClient, const long lCommandTimeout = -1);

	virtual long Execute(const _bstr_t &bstrSql, const long lCommandTimeout = -1, bool bPrepared = false);
	virtual long Execute(ADODB::_CommandPtr &pCommand, const long lCommandTimeout = -1);

private:
	ADOComm& operator=(const ADOComm &ado);
	ADOComm(const ADOComm &ado);
};





class CommandHelper
{
public:
	CommandHelper() { 
		Create(); 
	}

	virtual ~CommandHelper() {}

	CommandHelper(const CommandHelper &comm) {
		pCommand_ = comm.pCommand_;
	}

	CommandHelper& operator=(const CommandHelper &comm) {
		if (pCommand_ != comm.pCommand_) {
			pCommand_ = comm.pCommand_;
		}
		return *this;
	}

	ADODB::_CommandPtr& GetCommand() { 
		return pCommand_; 
	}

	operator ADODB::_CommandPtr&() { 
		return pCommand_; 
	}

	_bstr_t GetCommandText() const { 
		return pCommand_->GetCommandText(); 
	}

	ADODB::CommandTypeEnum GetCommandType() const {
		return pCommand_->GetCommandType();
	}

	void SetCommandText(const _bstr_t &bstrCmd, bool bPrepared = false, ADODB::CommandTypeEnum plCmdType = ADODB::adCmdText) {
		pCommand_->PutCommandText(bstrCmd);
		pCommand_->PutCommandType(plCmdType);
		pCommand_->PutPrepared(bPrepared ? VARIANT_TRUE : VARIANT_FALSE);
	}

	void SetCommandProc(const _bstr_t &bstrCmd, bool bPrepared = false) {
		SetCommandText(bstrCmd, bPrepared, ADODB::adCmdStoredProc);
	}

	void AddParamInput(const _bstr_t &bstrName, ADODB::DataTypeEnum Type, long lSize = 0, const _variant_t &vtValue = vtMissing, 
			ADODB::ParameterDirectionEnum Direction = ADODB::adParamInput) 
	{
		pCommand_->GetParameters()->Append(pCommand_->CreateParameter(bstrName, Type, Direction, lSize, vtValue));
	}

	void AddParamInputVarchar(const _bstr_t &bstrName, const _bstr_t &bsValue, long lSize = 4000) 
	{
		AddParamInput(bstrName, ADODB::adVarChar, lSize, bsValue);
	}

	void AddParamInputInt(const _bstr_t &bstrName, long lValue, long lSize = 0) 
	{
		AddParamInput(bstrName, ADODB::adInteger, lSize, lValue);
	}

	void AddParamInputFloat(const _bstr_t &bstrName, double dvalue, long lSize = 0) 
	{
		AddParamInput(bstrName, ADODB::adDouble, lSize, dvalue);
	}

	void AddParamReturnValue(const _bstr_t &bstrName) {
		AddParamInput(bstrName, ADODB::adInteger, 0, vtMissing, ADODB::adParamReturnValue);
	}

	void AddParamOutput(const _bstr_t &bstrName, ADODB::DataTypeEnum Type, long lSize = 0) {
		AddParamInput(bstrName, Type, lSize, vtMissing, ADODB::adParamOutput);
	}

	void AddParamInputOutput(const _bstr_t &bstrName, ADODB::DataTypeEnum Type, const _variant_t &vtValue, long lSize = 0) {
		AddParamInput(bstrName, Type, lSize, vtValue, ADODB::adParamInputOutput);
	}

	_variant_t GetParamValue(const _bstr_t &bstrName) const {
		return pCommand_->GetParameters()->GetItem(bstrName)->GetValue();
	}

	void SetParamValue(const _bstr_t &bstrName, const _variant_t &vtValue) {
		pCommand_->GetParameters()->GetItem(bstrName)->PutValue(vtValue);
	}


private:
	void Create() {
		pCommand_.CreateInstance(__uuidof(ADODB::Command)); 
	}

private:
	ADODB::_CommandPtr pCommand_;
};
	



//////////////////////////////////////////////////////////////////////
// ConnectionString class
template <class StrT>
class ConnectionStringT
{
public:
	ConnectionStringT(const StrT &strProvider, const StrT &strServerColumn = "Data Source") 
		: strProvider_(strProvider), strServerColumn_(strServerColumn) {}

	StrT operator()(const StrT &strServer) 
	{
		StrT Str = strProvider_;
		Str += ";";
		Str += strServerColumn_;
		Str += "=";
		Str += strServer;
		return Str;
	}

	StrT operator()(const StrT &strServer, const StrT &strUID, const StrT &strPWD) 
	{
		StrT strCStr = operator()(strServer);
		strCStr += ";User ID=";
		strCStr += strUID;
		strCStr += ";Password=";
		strCStr += strPWD;
		return strCStr;
	}

	StrT operator()(const StrT &strServer, const StrT &strUID, const StrT &strPWD, const StrT &strDatabase) 
	{
		StrT strCStr = operator()(strServer, strUID, strPWD);
		strCStr += ";Initial Catalog=";
		strCStr += strDatabase;
		return strCStr;
	}

private:
	StrT strProvider_;
	StrT strServerColumn_;
};




typedef ConnectionStringT<_bstr_t> ConnectionString;


// Provider
static ConnectionString ORAOLE("Provider=MSDAORA");
static ConnectionString SQLOLE("Provider=SQLOLEDB");
static ConnectionString MDBOLE("Provider=Microsoft.Jet.OLEDB.4.0");

static ConnectionString ODBC("Provider=MSDASQL", "DSN");
static ConnectionString SQLDriver("Provider=MSDASQL;Driver={SQL Server}", "Server");



} // namespace GLASS

#endif 


//////////////////////////////////////////////////////////////////////
// Copyright (c) 2004 by cdecl (byung-kyu kim)
// EMail : cdecl@interpark.com
//////////////////////////////////////////////////////////////////////









