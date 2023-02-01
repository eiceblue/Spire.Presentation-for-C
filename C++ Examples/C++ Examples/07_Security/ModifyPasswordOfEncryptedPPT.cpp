#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_4.pptx";
	std::wstring outputFile = OutputPath"ModifyPasswordOfEncryptedPPT.pptx";

	//Create a PowerPoint document.
	Presentation* presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str(), L"123456");

	//Remove the encryption.
	presentation->RemoveEncryption();

	//Protect the document by setting a new password.
	presentation->Protect(L"654321");

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;

}
