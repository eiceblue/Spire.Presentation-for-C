#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_4.pptx";
	std::wstring outputFile = OutputPath"CheckPasswordProtection.txt";

	//Create Presentation
	Presentation* presentation = new Presentation();

	//Check whether a PPT document is password protected
	bool isProtected = presentation->IsPasswordProtected(inputFile.c_str());

	wofstream outFile(outputFile, ios::out);
	outFile << "The file is " << (isProtected ? "password " : "not password ") << "protected!" << endl;

	//Save the file
	outFile.close();
	delete presentation;

}
