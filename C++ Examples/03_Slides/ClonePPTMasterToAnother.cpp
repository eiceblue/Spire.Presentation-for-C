#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile_1 = DataPath"CloneMaster1.pptx";
	std::wstring inputFile_2 = DataPath"CloneMaster2.pptx";
	std::wstring outputFile = OutputPath"ClonePPTMasterToAnother.pptx";

	//Load PPT1 from disk
	Presentation* presentation1 = new Presentation();
	presentation1->LoadFromFile(inputFile_1.c_str());

	//Load PPT2 from disk
	Presentation* presentation2 = new Presentation();
	presentation2->LoadFromFile(inputFile_2.c_str());

	//Add masters from PPT1 to PPT2
	for (int m = 0; m < presentation1->GetMasters()->GetCount(); m++)
	{
		IMasterSlide* masterSlide = presentation1->GetMasters()->GetItem(m);
		presentation2->GetMasters()->AppendSlide(masterSlide);
	}

	//Save the document
	presentation2->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation2;
}
