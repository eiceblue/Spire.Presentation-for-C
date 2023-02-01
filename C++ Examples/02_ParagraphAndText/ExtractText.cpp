#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ExtractText.pptx";
	std::wstring outputFile = OutputPath"ExtractText.txt";

	//Create a PPT document and load file
	Presentation* presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	wofstream desFile(outputFile, ios::out);
	//Foreach the slide and extract text
	for (int i = 0; i < presentation->GetSlides()->GetCount(); i++)
	{
		ISlide* slide = presentation->GetSlides()->GetItem(i);
		for (int s = 0; s < slide->GetShapes()->GetCount(); s++)
		{
			IShape* shape = slide->GetShapes()->GetItem(s);
			if (dynamic_cast<IAutoShape*>(shape) != nullptr)
			{
				for (int t = 0; t < (dynamic_cast<IAutoShape*>(shape))->GetTextFrame()->GetParagraphs()->GetCount(); t++)
				{
					TextParagraph* tp =
						(dynamic_cast<IAutoShape*>(shape))->GetTextFrame()->GetParagraphs()->GetItem(t);
					desFile << tp->GetText() << endl;
				}
			}
		}
	}
	desFile.close();
	delete presentation;

}
