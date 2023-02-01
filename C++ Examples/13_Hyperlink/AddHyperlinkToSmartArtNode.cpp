#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"SmartArtNode.pptx";
	std::wstring outputFile = OutputPath"AddHyperlinkToSmartArtNode.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first slide.
	ISlide* slide = ppt->GetSlides()->GetItem(0);

	///Get the smartArt shape
	ISmartArt* sr = dynamic_cast<ISmartArt*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));
	//Add hylerlinks to the nodes
	ISmartArtNode* node = sr->GetNodes()->GetItem(0);
	node->SetClick(new ClickHyperlink(ppt->GetSlides()->GetItem(1)));
	node = sr->GetNodes()->GetItem(1);
	node->SetClick(new ClickHyperlink(ppt->GetSlides()->GetItem(2)));
	node = sr->GetNodes()->GetItem(2);
	node->SetClick(new ClickHyperlink(ppt->GetSlides()->GetItem(3)));

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
