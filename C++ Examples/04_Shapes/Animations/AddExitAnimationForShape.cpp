#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile = OutputPath"AddExitAnimationForShape.pptx";

	//Create an instance of presentation document
	Presentation* ppt = new Presentation();

	//Get the first slide
	ISlide* slide = ppt->GetSlides()->GetItem(0);

	//Set background image
	std::wstring ImageFile = DataPath"bg.png";
	RectangleF* rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	slide->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	slide->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Add a shape to the slide
	IShape* starShape = slide->GetShapes()->AppendShape(ShapeType::FivePointedStar, new RectangleF(250, 100, 200, 200));
	starShape->GetFill()->SetFillType(FillFormatType::Solid);
	starShape->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::LightBlue);

	//Add random bars effect to the shape
	AnimationEffect* effect = slide->GetTimeline()->GetMainSequence()->AddEffect(starShape, AnimationEffectType::RandomBars);

	//Change effect type from entrance to exit
	effect->SetPresetClassType(TimeNodePresetClassType::Exit);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
