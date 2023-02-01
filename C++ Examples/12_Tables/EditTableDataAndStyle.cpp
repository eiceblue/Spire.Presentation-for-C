#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_1.pptx";
	std::wstring outputFile = OutputPath"EditTableDataAndStyle.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Store the data used in replacement in string [).
	std::vector<std::wstring> str = { L"Germany",L"Berlin",L"Europe",L"0152458",L"20860000" };

	ITable* table = nullptr;

	//Get the table in PowerPoint document.
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		IShape* shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);

		if (dynamic_cast<ITable*>(shape) != nullptr)
		{
			table = dynamic_cast<ITable*>(shape);

			//Change the style of table.
			table->SetStylePreset(TableStylePreset::LightStyle1Accent2);

			for (int i = 0; i < table->GetColumnsList()->GetCount(); i++)
			{
				//Replace the data in cell.
				table->GetItem(i, 2)->GetTextFrame()->SetText(str[i].c_str());

				//Set the highlightcolor.
				table->GetItem(i, 2)->GetTextFrame()->GetTextRange()->GetHighlightColor()->SetColor(Color::GetBlueViolet());
			}
		}
	}
	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
