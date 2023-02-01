#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"RemoveAllDigitalSignatures.pptx";
	std::wstring outputFile = OutputPath"RemoveAllDigitalSignatures.pptx";

	//Create a PowerPoint document.
	Presentation* ppt = new Presentation();

	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Remove all digital signatures
	if (ppt->GetIsDigitallySigned())
	{
		ppt->RemoveAllDigitalSignatures();
	}
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;

}
