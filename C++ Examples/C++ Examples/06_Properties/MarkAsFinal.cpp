#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"GetProperties.pptx";
	std::wstring outputFile = OutputPath"GetBuiltinProperties.txt";

	//Create PPT document
	Presentation* presentation = new Presentation();

	//Load the PPT document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Get the builtin properties 
	std::wstring application = presentation->GetDocumentProperty()->GetApplication();
	std::wstring author = presentation->GetDocumentProperty()->GetAuthor();
	std::wstring company = presentation->GetDocumentProperty()->GetCompany();
	std::wstring keywords = presentation->GetDocumentProperty()->GetKeywords();
	std::wstring comments = presentation->GetDocumentProperty()->GetComments();
	std::wstring category = presentation->GetDocumentProperty()->GetCategory();
	std::wstring title = presentation->GetDocumentProperty()->GetTitle();
	std::wstring subject = presentation->GetDocumentProperty()->GetSubject();

	wofstream outFile(outputFile, ios::out);
	outFile << "DocumentProperty.Application: " << application.c_str() << endl;
	outFile << "DocumentProperty.Author: " << author.c_str() << endl;
	outFile << "DocumentProperty.Company " << company.c_str() << endl;
	outFile << "DocumentProperty.Keywords: " << keywords.c_str() << endl;
	outFile << "DocumentProperty.Comments: " << comments.c_str() << endl;
	outFile << "DocumentProperty.Category: " << category.c_str() << endl;
	outFile << "DocumentProperty.Title: " << title.c_str() << endl;
	outFile << "DocumentProperty.Subject: " << subject.c_str();

	//Save them to a txt file
	outFile.close();
	delete presentation;

}
