#include "pch.h"
#include <codecvt>

using namespace Spire::Presentation;

string  wstring2string(const std::wstring& wstr)
{
	std::string result;
	result.reserve(wstr.size());
	for (size_t i = 0; i < wstr.size(); ++i)
	{
		result += static_cast<char>(wstr[i] & 0xFF);
	}
	return result;
}

int main()
{
	wstring inputFile = DATAPATH"ExtractText.pptx";
	wstring outputFile = OUTPUTPATH"ExtractText.txt";

	//Create a PPT document and load file
	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	wofstream desFile(wstring2string(outputFile));
	desFile.imbue(std::locale(desFile.getloc(), new std::codecvt_utf8<wchar_t>));

	//Foreach the slide and extract text
	for (int i = 0; i < presentation->GetSlides()->GetCount(); i++)
	{
		intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(i);
		for (int s = 0; s < slide->GetShapes()->GetCount(); s++)
		{
			intrusive_ptr<IShape> shape = slide->GetShapes()->GetItem(s);
			if (Object::CheckType<IAutoShape>(shape))
			{
				for (int t = 0; t < (Object::Dynamic_cast<IAutoShape>(shape))->GetTextFrame()->GetParagraphs()->GetCount(); t++)
				{
					cout << "IAutoShape" << endl;
					intrusive_ptr<TextParagraph> tp =
						(Object::Dynamic_cast<IAutoShape>(shape))->GetTextFrame()->GetParagraphs()->GetItem(t);
					desFile << tp->GetText() << endl;
				}
			}
		}
	}
	desFile.close();
}

