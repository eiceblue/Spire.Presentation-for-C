#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"SetTableStyle.pptx";
	std::wstring outputFile = OutputPath"SetTableStyle.pptx";

	//Creat a ppt document and load file
	Presentation* ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Get tbe table
	ITable* table = nullptr;
	for (int s = 0; s < ppt->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		IShape* shape = ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);

		if (dynamic_cast<ITable*>(shape) != nullptr)
		{
			table = dynamic_cast<ITable*>(shape);

			//Set the table style from TableStylePreset and apply it to selected table
			table->SetStylePreset(TableStylePreset::MediumStyle1Accent2);
		}
	}
	//Save the file
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;
}
