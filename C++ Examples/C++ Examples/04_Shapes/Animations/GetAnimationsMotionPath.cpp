#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

const wstring commentToString(MotionCommandPathType  comment)
{
	switch (comment)
	{
	case Spire::Presentation::MotionCommandPathType::MoveTo:
		return L"MoveTo";
		break;
	case Spire::Presentation::MotionCommandPathType::LineTo:
		return L"LineTo";
		break;
	case Spire::Presentation::MotionCommandPathType::CurveTo:
		return L"CurveTo";
		break;
	case Spire::Presentation::MotionCommandPathType::CloseLoop:
		return L"CloseLoop";
		break;
	case Spire::Presentation::MotionCommandPathType::End:
		return L"End";
		break;
	default:
		break;
	}
}

int main()
{
	std::wstring inputFile = DataPath"GetAnimationsMotionPath.pptx";
	std::wstring outputFile = OutputPath"GetAnimationsMotionPath.txt";

	Presentation* presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());
	ISlide* slide = presentation->GetSlides()->GetItem(0);
	//Get the first shape
	IShape* shape = slide->GetShapes()->GetItem(0);
	wofstream outFile(outputFile);
	int j = 1;
	//Traverse all animations

	for (int e = 0; e < slide->GetTimeline()->GetMainSequence()->GetCount(); e++)
	{
		AnimationEffect* effect = slide->GetTimeline()->GetMainSequence()->GetItem(e);
		int* p = effect->GetShapeTarget()->GetIntPtr();
		int* q = shape->GetIntPtr();
		if (*p == *q)
		{
			//Get MotionPath
			MotionPath* path = (dynamic_cast<AnimationMotion*>(effect->GetCommonBehaviorCollection()->GetItem(0)))->GetPath();

			//Get all points in the path
			for (int i = 0; i < path->GetCount(); i++)
			{
				MotionCmdPath* motionCmdPath = path->GetItem(i);
				std::vector<PointF*> points = motionCmdPath->GetPoints();
				MotionCommandPathType type = motionCmdPath->GetCommandType();
				if (points.size() > 0)
				{
					for (auto point : points)
					{
						outFile << std::to_string(i + 1).c_str() << "  MotionType: " << commentToString(type) << " -> X: " << std::to_string(point->GetX()).c_str() << ", Y: " << std::to_string(point->GetY()).c_str() << endl;
					}
					j++;
				}
			}
		}
	}
	delete presentation;

}
