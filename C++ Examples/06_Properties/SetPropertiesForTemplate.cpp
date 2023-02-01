#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile_pptx = OutputPath"SetPropertiesForTemplate.pptx";
	std::wstring outputFile_ppt = OutputPath"SetPropertiesForTemplate.ppt";
	std::wstring outputFile_odp = OutputPath"SetPropertiesForTemplate.odp";

	//Create a document
	Presentation* presentation = new Presentation();

	//Set the DocumentProperty 
	presentation->GetDocumentProperty()->SetApplication(L"Spire.Presentation");
	presentation->GetDocumentProperty()->SetAuthor(L"E-iceblue");
	presentation->GetDocumentProperty()->SetCompany(L"E-iceblue Co., Ltd.");
	presentation->GetDocumentProperty()->SetKeywords(L"Demo File");
	presentation->GetDocumentProperty()->SetComments(L"This file is used to test Spire.Presentation.");
	presentation->GetDocumentProperty()->SetCategory(L"Demo");
	presentation->GetDocumentProperty()->SetTitle(L"This is a demo file.");
	presentation->GetDocumentProperty()->SetSubject(L"Test");

	//Create the .pptx template
	presentation->SaveToFile(outputFile_pptx.c_str(), FileFormat::Pptx2013);

	//Create the .odp template
	presentation->SaveToFile(outputFile_odp.c_str(), FileFormat::ODP);

	//Create the .ppt template
	presentation->SaveToFile(outputFile_ppt.c_str(), FileFormat::PPT);
	delete presentation;

}
