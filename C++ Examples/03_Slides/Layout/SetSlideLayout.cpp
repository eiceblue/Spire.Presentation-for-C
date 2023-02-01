#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile = OutputPath"SetSlideLayout.pptx";

	//Create an instance of presentation document
	Presentation* ppt = new Presentation();

	//Remove the first slide
	ppt->GetSlides()->RemoveAt(0);

	//Append a slide and set the layout for slide
	ISlide* slide = ppt->GetSlides()->Append(SlideLayoutType::Title);

	//Add content for Title and Text
	IAutoShape* shape = dynamic_cast<IAutoShape*>(slide->GetShapes()->GetItem(0));
	shape->GetTextFrame()->SetText(L"Hello Wolrd! -> This is title");

	shape = dynamic_cast<IAutoShape*>(slide->GetShapes()->GetItem(1));
	shape->GetTextFrame()->SetText(L"E-iceblue Support Team -> This is content");

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
