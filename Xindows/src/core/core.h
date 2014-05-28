
#ifndef __XINDOWS_CORE_H__
#define __XINDOWS_CORE_H__

#pragma warning ( disable : 4200 )  // nonstandard extension used : zero-sized array in struct/union
#pragma warning ( disable : 4244 )  // conversion from XXX to YYY, possible loss of data
#pragma warning ( disable : 4407 )  // cast between different pointer to member representations, compiler may generate incorrect code
#pragma warning ( disable : 4996 )  // function or variable may be unsafe.

#include <ComCat.h>
#include <WindowsX.h>
#include <Shlwapi.h>
#include <CommCtrl.h>
#include <CommDlg.h>
#include <Ole2.h>
#include <OleCtl.h>
#include <WinInet.h>
#include <DispEx.h>
#include <DocObj.h>

#include "../Xindows_h.h"

#include "include/Xindowsdid.h"

#include "../XindowsUtil/src/XindowsUtil.h"

#include "base/IPrivateUnknown.h"
#include "base/CoreRc.h"
#include "base/BaseDef.h"
#include "base/GlobalData.h"

#include "util/ColorUtil.h"
#include "util/Color3D.h"
#include "util/ColorInfo.h"
#include "util/ColorValue.h"
#include "util/ComUtil.h"
#include "util/Himetric.h"
#include "util/MiscUtil.h"
#include "util/FormsChar.h"
#include "util/FormsString.h"
#include "util/Transform.h"
#include "util/UnitValue.h"
#include "util/OffScreenContext.h"
#include "util/BrushUtil.h"
#include "util/ButtonUtil.h"
#include "util/WindowsUtil.h"
#include "util/ShapeUtil.h"

#include "com/Tearoff.h"
#include "com/ErrorInfo.h"
#include "com/ExcepInfo.h"
#include "com/DispParams.h"
#include "com/Connect.h"

#include "base/ThreadState.h"
#include "base/Dll.h"
#include "base/Timer.h"
#include "base/Task.h"
#include "base/Base.h"
#include "base/Inplcae.h"
#include "base/Server.h"
#include "base/Tooltip.h"
#include "base/DownBase.h"
#include "base/TimerCtx.h"

#include "gen/funcsig.h"

#endif //__XINDOWS_CORE_H__