#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"AddSection.pptx";
	std::wstring outputFile = OutputPath"GetSectionIndex.txt";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	Section* section = ppt->GetSectionList()->GetItem(0);

	int index = ppt->GetSectionList()->IndexOf(section);

	//Save to file.
	std::wofstream out;
	out.open(outputFile);
	out.flush();
	out << L"index:" + std::to_wstring(index);
	out.close();
	delete ppt;

}
