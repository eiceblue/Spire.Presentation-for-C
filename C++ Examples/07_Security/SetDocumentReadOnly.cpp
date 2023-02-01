#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"SetDocumentReadOnly.pptx";
	std::wstring outputFile = OutputPath"SetDocumentReadOnly.pptx";

	//Load a PPT document
	Presentation* presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Get the password that the user entered
	std::wstring password = L"e-iceblue";

	//Protect the document with the password
	presentation->Protect(password.c_str());

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;

}
