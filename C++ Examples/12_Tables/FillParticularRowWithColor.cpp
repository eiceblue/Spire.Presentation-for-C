#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_1.pptx";
	std::wstring outputFile = OutputPath"FillParticularRowWithColor.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();
	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Fill particular table row with color.
	ITable* table = nullptr;
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		IShape* shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		if (dynamic_cast<ITable*>(shape) != nullptr)
		{
			table = dynamic_cast<ITable*>(shape);

			TableRow* row = table->GetTableRows()->GetItem(1);
			for (int n = 0; n < row->GetCount(); n++)
			{
				Cell* cell = row->GetItem(n);
				cell->GetFillFormat()->SetFillType(FillFormatType::Solid);
				cell->GetFillFormat()->GetSolidColor()->SetColor(Color::GetPink());
			}
		}
	}
	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
