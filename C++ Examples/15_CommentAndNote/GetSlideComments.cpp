#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Comments.pptx";
	std::wstring outputFile = OutputPath"GetSlideComments.txt";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	wofstream outFile(outputFile, ios::out);

	//Loop through comments
	for (int i = 0; i < ppt->GetCommentAuthors()->GetCount(); i++)
	{
		ICommentAuthor* commentAuthor = ppt->GetCommentAuthors()->GetItem(i);
		for (int j = 0; j < commentAuthor->GetCommentsList()->GetCount(); j++)
		{
			//Get comment information
			Comment* comment = commentAuthor->GetCommentsList()->GetItem(j);

			outFile << "Comment text : " << comment->GetText() << endl;
			outFile << "Comment author : " << comment->GetAuthorName() << endl;
			outFile << "Posted on time : " << comment->GetDateTime()->ToString() << endl;
		}
	}
	outFile.close();
	delete ppt;
}
