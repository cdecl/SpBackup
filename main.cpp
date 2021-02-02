
#include <windows.h>
#include <shlwapi.h>
#pragma comment(lib, "shlwapi.lib")

#include <iostream>
#include <string>
#include <sstream>
#include <fstream>
using namespace std;

#include "ADO.h"

ostream& operator<<(ostream& os, const _bstr_t &bs)
{
	LPCSTR pstr = (LPCSTR)bs;
	if (pstr) {
		os << pstr << endl;
	}

	return os;
}

pair<string, string> split_kv(string str, const string& fd)
{
	string v;
	string k = str;
	auto pos = str.find(fd);
	if (pos != string::npos) {
		k = str.substr(0, pos);
		v = str.substr(pos + fd.size());
	}
	return make_pair(k, v);
}

string replaceall(string str, const string &strFind, const string &strReplaced) 
{
	string::size_type p = 0;	

	while (true) {
		p = str.find(strFind, p);

		if (p == string::npos) {
			break;
		}

		str = str.replace(p, strFind.length(), strReplaced);
		p += strReplaced.length();
	}

	return str;
}

//string GetTypeName(const string& strType)
//{
//	string strName;
//
//	if (strType == "FN") {
//		strName = "ScalarFunction";
//	}
//	else if (strType == "IF") {
//		strName = "InlineTableFunction";
//	}
//	else if (strType == "P") {
//		strName = "Procedure";
//	}
//	else if (strType == "TF") {
//		strName = "TableFunction";
//	}
//	else if (strType == "TR") {
//		strName = "Trigger";
//	}
//	else if (strType == "V") {
//		strName = "View";
//	}
//	
//	return strName;
//}

string GetTextObject(GLASS::ADOComm &adoProc, const string &strProcName, const string &strType)
{
	GLASS::CommandHelper cmd;
	auto [k, v] = split_kv(strType, ":");

	if (v.length() > 0) {
		string sProcName = v;
		sProcName += " ?";

		cmd.SetCommandText(sProcName.c_str());
		cmd.AddParamInputVarchar("@obj_name", strProcName.c_str(), 200);
	}
	else {
		cmd.SetCommandText("SELECT text AS txt FROM syscomments with(nolock) WHERE id = object_id( ? )");
		cmd.AddParamInputVarchar("@obj_name", strProcName.c_str(), 200);
	}

	adoProc.OpenRs(cmd);

	ostringstream os;
	string str;

	while (!adoProc.IsEOF()) {
		str = (LPCSTR)(_bstr_t)adoProc(0L);
		str = replaceall(str, "\r\n", "\n");

		os << str << '\n';
		adoProc.MoveNext();
	}

	return os.str();
}

void ProcedureBackup(GLASS::ADOComm &adoProc, const string &strPath, const string &strProcName, const string& strType)
{
	try {
		string strFilePath = strPath;

		if (strFilePath[strFilePath.length() - 1] != '\\') {
			strFilePath += '\\';
		}
		strFilePath += strProcName;
		strFilePath += ".sql";

		string str = GetTextObject(adoProc, strProcName, strType);

		ofstream fout(strFilePath.c_str());
		fout << str << endl;
		fout.close();

	}
	catch (exception &ex) {
		cerr << ex.what() << endl;
	}

	adoProc.CloseRs();
}



void Run(const string &strConnectionString, const string &strPath, const string &strType)
{
	bool bDirectory = PathIsDirectory(strPath.c_str()) ? true : false;

	ofstream fout;

	if (!bDirectory) {
		fout.open(strPath.c_str());
	}

	std::ostringstream osSql;
	auto [k, v] = split_kv(strType, ":");

	string sql = R"(
		SELECT	name 
		FROM	sysobjects with(nolock)
		WHERE	uid = 1
		AND		Type = ?
		ORDER BY name
	)";
		
	GLASS::CommandHelper cmd;
	cmd.SetCommandText(sql.c_str());
	cmd.AddParamInputVarchar("@obj_name", k.c_str(), 8);

	GLASS::ADOComm ado;
	ado.Create(strConnectionString.c_str());
	ado.OpenRs(cmd);

	const long lRowCount = ado.GetRecordCount();
	long i = 0;

	string strName;

	GLASS::ADOComm adoProc;
	adoProc.Create(strConnectionString.c_str());

	while (!ado.IsEOF()) {
		strName = (LPCSTR)(_bstr_t)ado("name");

		if (bDirectory) {
			ProcedureBackup(adoProc, strPath, strName, strType);
		}
		else {
			fout << GetTextObject(adoProc, strName, strType) << '\n';
			fout << '\n';
			fout << "GO" << '\n';
			fout << "--** END : " << strName << '\n';
			fout << "--**************************************************************" << '\n';
			fout << endl;
		}

		cout << '\r' << string(75, ' ');
		cout << '\r';
		cout << "[" << ++i << "/" << lRowCount << "] " << strName;

		ado.MoveNext();
	}
}


void Usage()
{
	cout << "Usage: SpBackup.exe [connectionString] [path] [type] " << endl;
	cout << "  - connectionString : DB 연결문자열 " << endl;
	cout << "  - path : 저장폴더(객체명으로 각각 생성) 혹은 저장파일(지정한 파일생성) " << endl;
	cout << "  - type : 객체타입  " << endl;
	cout << "      P = 저장 프로시저 (default)" << endl;
	cout << "      FN = 스칼라 함수" << endl;
	cout << "      IF = 인라인 테이블 함수" << endl;
	cout << "      TF = 테이블 함수" << endl;
	cout << "      TR = 트리거" << endl;
	cout << "      V = 뷰" << endl;
	cout << "      U:<GetDDLProcName> = 테이블 (U:TableCreateScriptProcedure)" << endl;
	cout << "  - verion: 210202.2 " << "(_MSC_VER:" << _MSC_VER << ")" << endl;
}

int main(int argc, char *argv[])
{
	if (argc < 3) {
		Usage();
		return 1;
	}

	::CoInitialize(NULL);

	try {
		string strType = "P";

		if (argc > 3) {
			strType = argv[3];
		}
		
		Run(argv[1], argv[2], strType);
	}
	catch (_com_error &e) {
		cout << e.Description() << endl;
	}

	::CoUninitialize();
}
