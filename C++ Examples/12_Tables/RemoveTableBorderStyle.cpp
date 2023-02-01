#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_1.pptx";
	std::wstring outputFile = OutputPath"RemoveTableBorderStyle.pptx";

	//Create a PowerPoint document.
	Presentation* presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

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
						cell->GetBorderTop()->SetFillType(FillFormatType::None);
						cell->GetBorderBottom()->SetFillType(FillFormatType::None);
						cell->GetBorderLeft()->SetFillType(FillFormatType::None);
						cell->GetBorderRight()->SetFillType(FillFormatType::None);
					}
				}
			}
		}
	}
	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
