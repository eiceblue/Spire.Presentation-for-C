#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile = OutputPath"CustomPathAnimation.pptx";

	//Create PPT document
	Presentation* ppt = new Presentation();

	//Add shape
	IAutoShape* shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(0, 0, 200, 200));

	//Add animation
	AnimationEffect* effect = ppt->GetSlides()->GetItem(0)->GetTimeline()->GetMainSequence()->AddEffect(shape, AnimationEffectType::PathUser);
	CommonBehaviorCollection* common = effect->GetCommonBehaviorCollection();
	AnimationMotion* motion = dynamic_cast<AnimationMotion*>(common->GetItem(0));
	motion->SetOrigin(AnimationMotionOrigin::Layout);
	motion->SetPathEditMode(AnimationMotionPathEditMode::Relative);

	//Add moin path
	MotionPath* moinPath = new MotionPath();
	moinPath->Add(MotionCommandPathType::MoveTo, { new PointF(0, 0) }, MotionPathPointsType::CurveAuto, true);
	moinPath->Add(MotionCommandPathType::LineTo, { new PointF(0.1f, 0.1f) }, MotionPathPointsType::CurveAuto, true);
	moinPath->Add(MotionCommandPathType::LineTo, { new PointF(-0.1f, 0.2f) }, MotionPathPointsType::CurveAuto, true);
	moinPath->Add(MotionCommandPathType::End, {}, MotionPathPointsType::CurveStraight, true);
	motion->SetPath(moinPath);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;

}
