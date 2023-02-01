#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_1.pptx";
	std::wstring outputFile = OutputPath"TraverseThroughCells.txt";

	//Create a PowerPonit document.
	Presentation* presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	wofstream content(outputFile);
	content << "The data in cells of this PowerPoint file is: " << endl;

	//Get the table.
	ITable* table = nullptr;
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		IShape* shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		if (dynamic_cast<ITable*>(shape) != nullptr)
		{
			table = dynamic_cast<ITable*>(shape);

			//Traverse through the cells of table.
			for (int r = 0; r < table->GetTableRows()->GetCount(); r++)
			{
				TableRow* row = table->GetTableRows()->GetItem(r);
				for (int c = 0; c < row->GetCount(); c++)
				{
					Cell* cell = row->GetItem(c);
					content << cell->GetTextFrame()->GetText() << endl;
				}
				content << endl;
			}
		}
	}
	//Save to file.
	content.close();
	delete presentation;
}
