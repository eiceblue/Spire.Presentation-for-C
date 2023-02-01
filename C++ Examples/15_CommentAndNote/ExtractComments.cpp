#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_5.pptx";
	std::wstring outputFile = OutputPath"ExtractComments.txt";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	wofstream outFile(outputFile, ios::out);

	//Get all comments from the first slide.
	std::vector<Comment*> comments = ppt->GetSlides()->GetItem(0)->GetComments();

	//Save the comments in txt file.
	for (int i = 0; i < comments.size(); i++)
	{
		outFile << comments[i]->GetText() << endl;
	}
	outFile.close();
	delete ppt;
}
