#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"HeaderAndFooter.pptx";
	std::wstring outputFile = OutputPath"HeaderAndFooter.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Add footer
	ppt->SetFooterText(L"Demo of Spire.Presentation");

	//Set the footer visible
	ppt->SetFooterVisible(true);

	//Set the page number visible
	ppt->SetSlideNumberVisible(true);

	//Set the date visible
	ppt->SetDateTimeVisible(true);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;

}
