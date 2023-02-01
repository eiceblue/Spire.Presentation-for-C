#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Macros.ppt";
	std::wstring outputFile = OutputPath"RemoveVBAMacros.ppt";

	Presentation* presentation = new Presentation();

	//Load PPT file from disk
	presentation->LoadFromFile(inputFile.c_str());
	//Remove macros
	//Note, at present it only can work on macros in PPT file, has not supported for PPTM file yet.
	presentation->DeleteMacros();
	presentation->SaveToFile(outputFile.c_str(), FileFormat::PPT);
	delete presentation;
}
