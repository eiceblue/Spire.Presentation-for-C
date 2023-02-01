#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Bullets.pptx";
	std::wstring outputFile = OutputPath"PictureCustomBulletStyle.pptx";

	//Create an instance of presentation document
	Presentation* ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	//Get the second shape on the first slide
	IAutoShape* shape = dynamic_cast<IAutoShape*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(1));

	//Traverse through the paragraphs in the shape
	for (int t = 0; t < shape->GetTextFrame()->GetParagraphs()->GetCount(); t++) {
		TextParagraph* paragraph = shape->GetTextFrame()->GetParagraphs()->GetItem(t);
		//Set the bullet style of paragraph as picture
		paragraph->SetBulletType(TextBulletType::Picture);
		//Load a picture
		std::wstring inputImg = DataPath"icon.png";
		std::ifstream inputf(inputImg.c_str(), std::ios::in | std::ios::binary);
		Stream* stream = new Stream(inputf);
		paragraph->GetBulletPicture()->SetEmbedImage(ppt->GetImages()->Append(stream));
		stream->Close();
	}
	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
