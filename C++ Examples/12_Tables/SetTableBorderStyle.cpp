#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_1.pptx";
	std::wstring outputFile = OutputPath"SetTableBorderStyle.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Find the table by looping through all the slides, and then set borders for it. 
	for (int l = 0; l < presentation->GetSlides()->GetCount(); l++)
	{
		ISlide* slide = presentation->GetSlides()->GetItem(l);
		for (int s = 0; s < slide->GetShapes()->GetCount(); s++)
		{
			IShape* shape = slide->GetShapes()->GetItem(s);
			if (dynamic_cast<ITable*>(shape) != nullptr)
			{
				ITable* table = dynamic_cast<ITable*>(shape);
				for (int i = 0; i < table->GetTableRows()->GetCount(); i++)
				{
					TableRow* row = table->GetTableRows()->GetItem(i);
					for (int j = 0; j < row->GetCount(); j++)
					{
						Cell* cell = row->GetItem(j);
						cell->GetBorderTop()->SetFillType(FillFormatType::Solid);
						cell->GetBorderBottom()->SetFillType(FillFormatType::Solid);
						cell->GetBorderLeft()->SetFillType(FillFormatType::Solid);
						cell->GetBorderRight()->SetFillType(FillFormatType::Solid);
					}
				}
			}
		}
	}
	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
