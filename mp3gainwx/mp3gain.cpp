#include "wx/wxprec.h"

#ifndef WX_PRECOMP
#include "wx/wx.h"
#endif

#include "mainFrame.h"
class MyApp: public wxApp {
	virtual bool OnInit();
};


IMPLEMENT_APP(MyApp)

bool MyApp::OnInit() {
	mainFrame *frame = new mainFrame(NULL, ID_WINDOW_MAIN_FRAME, wxT("MP3Gain"));
	frame->Show(TRUE);
	SetTopWindow(frame);
	return TRUE;
};
