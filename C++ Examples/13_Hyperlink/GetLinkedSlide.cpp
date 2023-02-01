#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"linkedSlide.pptx";
	std::wstring outputFile = OutputPath"GetLinkedSlide.txt";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the second slide
	ISlide* slide = ppt->GetSlides()->GetItem(1);

	//Get the first shape of the second slide
	IAutoShape* shape = dynamic_cast<IAutoShape*>(slide->GetShapes()->GetItem(0));
	wofstream outFile(outputFile);
	//Get the linked slide index
	if (shape->GetClick()->GetActionType() == HyperlinkActionType::GotoSlide)
	{
		ISlide* targetSlide = shape->GetClick()->GetTargetSlide();
		outFile << "Linked slide number = " << targetSlide->GetSlideNumber() << endl;
	}
	outFile.close();
	delete ppt;

}
