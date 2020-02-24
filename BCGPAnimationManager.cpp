//*******************************************************************************
// COPYRIGHT NOTES
// ---------------
// This is a part of BCGControlBar Library Professional Edition
// Copyright (C) 1998-2018 BCGSoft Ltd.
// All rights reserved.
//
// This source code can be used, distributed or modified
// only under terms and conditions 
// of the accompanying license agreement.
//*******************************************************************************
//
// BCGPAnimationManager.cpp : implementation file
//

#include "stdafx.h"
#include <math.h>
#include "BCGPAnimationManager.h"
#include "bcgglobals.h"

#ifndef IID_PPV_ARGS

extern "C++"
{
    template<typename T> void** IID_PPV_ARGS_Helper(T** pp) 
    {
        static_cast<IUnknown*>(*pp);    // make sure everyone derives from IUnknown
        return reinterpret_cast<void**>(pp);
    }
}

#define IID_PPV_ARGS(ppType) __uuidof(**(ppType)), IID_PPV_ARGS_Helper(ppType)

#endif

#ifndef __UIAnimation_h__

interface IUIAnimationTransition;
interface IUIAnimationStoryboard;
interface IUIAnimationManagerEventHandler;
interface IUIAnimationPriorityComparison;
interface IUIAnimationTimerUpdateHandler;
interface IUIAnimationTimerEventHandler;
interface IUIAnimationVariableChangeHandler;
interface IUIAnimationVariableIntegerChangeHandler;
interface IUIAnimationStoryboardEventHandler;
interface IUIAnimationTimerClientEventHandler;

typedef DOUBLE UI_ANIMATION_SECONDS;

typedef enum
{	UI_ANIMATION_UPDATE_NO_CHANGE	= 0,
	UI_ANIMATION_UPDATE_VARIABLES_CHANGED	= 1
} 	
UI_ANIMATION_UPDATE_RESULT;

typedef enum
{	UI_ANIMATION_MANAGER_IDLE	= 0,
	UI_ANIMATION_MANAGER_BUSY	= 1
} 	
UI_ANIMATION_MANAGER_STATUS;

typedef enum 
{	UI_ANIMATION_MODE_DISABLED	= 0,
	UI_ANIMATION_MODE_SYSTEM_DEFAULT	= 1,
	UI_ANIMATION_MODE_ENABLED	= 2
} 	
UI_ANIMATION_MODE;

typedef enum 
{	UI_ANIMATION_IDLE_BEHAVIOR_CONTINUE	= 0,
	UI_ANIMATION_IDLE_BEHAVIOR_DISABLE	= 1
} 	
UI_ANIMATION_IDLE_BEHAVIOR;

typedef enum
{	UI_ANIMATION_SLOPE_INCREASING	= 0,
	UI_ANIMATION_SLOPE_DECREASING	= 1
} 	
UI_ANIMATION_SLOPE;

typedef enum
{	UI_ANIMATION_ROUNDING_NEAREST	= 0,
	UI_ANIMATION_ROUNDING_FLOOR	= 1,
	UI_ANIMATION_ROUNDING_CEILING	= 2
} 	
UI_ANIMATION_ROUNDING_MODE;

typedef struct __MIDL___MIDL_itf_UIAnimation_0000_0002_0003
{
    int _;
} 	*UI_ANIMATION_KEYFRAME;


typedef /* [public][public][v1_enum] */ 
enum __MIDL___MIDL_itf_UIAnimation_0000_0002_0002
{	UI_ANIMATION_SCHEDULING_UNEXPECTED_FAILURE	= 0,
	UI_ANIMATION_SCHEDULING_INSUFFICIENT_PRIORITY	= 1,
	UI_ANIMATION_SCHEDULING_ALREADY_SCHEDULED	= 2,
	UI_ANIMATION_SCHEDULING_SUCCEEDED	= 3,
	UI_ANIMATION_SCHEDULING_DEFERRED	= 4
} 	
UI_ANIMATION_SCHEDULING_RESULT;

typedef /* [public][public][public][public][v1_enum] */ 
enum __MIDL___MIDL_itf_UIAnimation_0000_0002_0001
{	UI_ANIMATION_STORYBOARD_BUILDING	= 0,
	UI_ANIMATION_STORYBOARD_SCHEDULED	= 1,
	UI_ANIMATION_STORYBOARD_CANCELLED	= 2,
	UI_ANIMATION_STORYBOARD_PLAYING	= 3,
	UI_ANIMATION_STORYBOARD_TRUNCATED	= 4,
	UI_ANIMATION_STORYBOARD_FINISHED	= 5,
	UI_ANIMATION_STORYBOARD_READY	= 6,
	UI_ANIMATION_STORYBOARD_INSUFFICIENT_PRIORITY	= 7
} 	UI_ANIMATION_STORYBOARD_STATUS;


EXTERN_C const IID IID_IUIAnimationManager;

MIDL_INTERFACE("9169896C-AC8D-4e7d-94E5-67FA4DC2F2E8")
IUIAnimationManager : public IUnknown
{
public:
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE CreateAnimationVariable( 
        /* [annotation][in] */ 
        DOUBLE initialValue,
        /* [annotation][retval][out] */ 
        IUIAnimationVariable **variable) = 0;
    
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE ScheduleTransition( 
        /* [annotation][in] */ 
        IUIAnimationVariable *variable,
        /* [annotation][in] */ 
        IUIAnimationTransition *transition,
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS timeNow) = 0;
    
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE CreateStoryboard( 
        /* [annotation][retval][out] */ 
        IUIAnimationStoryboard **storyboard) = 0;
    
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE FinishAllStoryboards( 
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS completionDeadline) = 0;
    
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE AbandonAllStoryboards( void) = 0;
    
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE Update( 
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS timeNow,
        /* [annotation][defaultvalue][out] */ 
        UI_ANIMATION_UPDATE_RESULT *updateResult = 0) = 0;
    
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE GetVariableFromTag( 
        /* [annotation][unique][in] */ 
        IUnknown *object,
        /* [annotation][in] */ 
        UINT32 id,
        /* [annotation][retval][out] */ 
        IUIAnimationVariable **variable) = 0;
    
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE GetStoryboardFromTag( 
        /* [annotation][unique][in] */ 
        IUnknown *object,
        /* [annotation][in] */ 
        UINT32 id,
        /* [annotation][retval][out] */ 
        IUIAnimationStoryboard **storyboard) = 0;
    
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE GetStatus( 
        /* [annotation][retval][out] */ 
        UI_ANIMATION_MANAGER_STATUS *status) = 0;
    
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE SetAnimationMode( 
        /* [annotation][in] */ 
        UI_ANIMATION_MODE mode) = 0;
    
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE Pause( void) = 0;
    
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE Resume( void) = 0;
    
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE SetManagerEventHandler( 
        /* [annotation][unique][in] */ 
        IUIAnimationManagerEventHandler *handler) = 0;
    
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE SetCancelPriorityComparison( 
        /* [annotation][unique][in] */ 
        IUIAnimationPriorityComparison *comparison) = 0;
    
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE SetTrimPriorityComparison( 
        /* [annotation][unique][in] */ 
        IUIAnimationPriorityComparison *comparison) = 0;
    
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE SetCompressPriorityComparison( 
        /* [annotation][unique][in] */ 
        IUIAnimationPriorityComparison *comparison) = 0;
    
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE SetConcludePriorityComparison( 
        /* [annotation][unique][in] */ 
        IUIAnimationPriorityComparison *comparison) = 0;
    
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE SetDefaultLongestAcceptableDelay( 
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS delay) = 0;
    
    virtual /* [annotation] */ 
    HRESULT STDMETHODCALLTYPE Shutdown( void) = 0;
    
};

MIDL_INTERFACE("6B0EFAD1-A053-41d6-9085-33A689144665")
IUIAnimationTimer : public IUnknown
{
public:
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE SetTimerUpdateHandler( 
        /* [annotation][unique][in] */ 
        IUIAnimationTimerUpdateHandler *updateHandler,
        /* [annotation][in] */ 
        UI_ANIMATION_IDLE_BEHAVIOR idleBehavior) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE SetTimerEventHandler( 
        /* [annotation][unique][in] */ 
        IUIAnimationTimerEventHandler *handler) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE Enable( void) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE Disable( void) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE IsEnabled( void) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE GetTime( 
        /* [annotation][out] */ 
        UI_ANIMATION_SECONDS *seconds) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE SetFrameRateThreshold( 
        /* [annotation][in] */ 
        UINT32 framesPerSecond) = 0;
};

MIDL_INTERFACE("CA5A14B1-D24F-48b8-8FE4-C78169BA954E")
IUIAnimationTransitionLibrary : public IUnknown
{
public:
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE CreateInstantaneousTransition( 
        /* [annotation][in] */ 
        DOUBLE finalValue,
        /* [annotation][retval][out] */ 
        IUIAnimationTransition **transition) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE CreateConstantTransition( 
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS duration,
        /* [annotation][retval][out] */ 
        IUIAnimationTransition **transition) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE CreateDiscreteTransition( 
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS delay,
        /* [annotation][in] */ 
        DOUBLE finalValue,
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS hold,
        /* [annotation][retval][out] */ 
        IUIAnimationTransition **transition) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE CreateLinearTransition( 
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS duration,
        /* [annotation][in] */ 
        DOUBLE finalValue,
        /* [annotation][retval][out] */ 
        IUIAnimationTransition **transition) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE CreateLinearTransitionFromSpeed( 
        /* [annotation][in] */ 
        DOUBLE speed,
        /* [annotation][in] */ 
        DOUBLE finalValue,
        /* [annotation][retval][out] */ 
        IUIAnimationTransition **transition) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE CreateSinusoidalTransitionFromVelocity( 
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS duration,
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS period,
        /* [annotation][retval][out] */ 
        IUIAnimationTransition **transition) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE CreateSinusoidalTransitionFromRange( 
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS duration,
        /* [annotation][in] */ 
        DOUBLE minimumValue,
        /* [annotation][in] */ 
        DOUBLE maximumValue,
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS period,
        /* [annotation][in] */ 
        UI_ANIMATION_SLOPE slope,
        /* [annotation][retval][out] */ 
        IUIAnimationTransition **transition) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE CreateAccelerateDecelerateTransition( 
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS duration,
        /* [annotation][in] */ 
        DOUBLE finalValue,
        /* [annotation][in] */ 
        DOUBLE accelerationRatio,
        /* [annotation][in] */ 
        DOUBLE decelerationRatio,
        /* [annotation][retval][out] */ 
        IUIAnimationTransition **transition) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE CreateReversalTransition( 
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS duration,
        /* [annotation][retval][out] */ 
        IUIAnimationTransition **transition) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE CreateCubicTransition( 
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS duration,
        /* [annotation][in] */ 
        DOUBLE finalValue,
        /* [annotation][in] */ 
        DOUBLE finalVelocity,
        /* [annotation][retval][out] */ 
        IUIAnimationTransition **transition) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE CreateSmoothStopTransition( 
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS maximumDuration,
        /* [annotation][in] */ 
        DOUBLE finalValue,
        /* [annotation][retval][out] */ 
        IUIAnimationTransition **transition) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE CreateParabolicTransitionFromAcceleration( 
        /* [annotation][in] */ 
        DOUBLE finalValue,
        /* [annotation][in] */ 
        DOUBLE finalVelocity,
        /* [annotation][in] */ 
        DOUBLE acceleration,
        /* [annotation][retval][out] */ 
        IUIAnimationTransition **transition) = 0;
    
};

MIDL_INTERFACE("8CEEB155-2849-4ce5-9448-91FF70E1E4D9")
IUIAnimationVariable : public IUnknown
{
public:
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE GetValue( 
        /* [annotation][retval][out] */ 
        DOUBLE *value) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE GetFinalValue( 
        /* [annotation][retval][out] */ 
        DOUBLE *finalValue) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE GetPreviousValue( 
        /* [annotation][retval][out] */ 
        DOUBLE *previousValue) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE GetIntegerValue( 
        /* [annotation][retval][out] */ 
        INT32 *value) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE GetFinalIntegerValue( 
        /* [annotation][retval][out] */ 
        INT32 *finalValue) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE GetPreviousIntegerValue( 
        /* [annotation][retval][out] */ 
        INT32 *previousValue) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE GetCurrentStoryboard( 
        /* [annotation][retval][out] */ 
        IUIAnimationStoryboard **storyboard) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE SetLowerBound( 
        /* [annotation][in] */ 
        DOUBLE bound) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE SetUpperBound( 
        /* [annotation][in] */ 
        DOUBLE bound) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE SetRoundingMode( 
        /* [annotation][in] */ 
        UI_ANIMATION_ROUNDING_MODE mode) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE SetTag( 
        /* [annotation][unique][in] */ 
        IUnknown *object,
        /* [annotation][in] */ 
        UINT32 id) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE GetTag( 
        /* [annotation][out] */ 
        IUnknown **object,
        /* [annotation][out] */ 
        UINT32 *id) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE SetVariableChangeHandler( 
        /* [annotation][unique][in] */ 
        IUIAnimationVariableChangeHandler *handler) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE SetVariableIntegerChangeHandler( 
        /* [annotation][unique][in] */ 
        IUIAnimationVariableIntegerChangeHandler *handler) = 0;
        
};

MIDL_INTERFACE("A8FF128F-9BF9-4af1-9E67-E5E410DEFB84")
IUIAnimationStoryboard : public IUnknown
{
public:
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE AddTransition( 
        /* [annotation][in] */ 
        IUIAnimationVariable *variable,
        /* [annotation][in] */ 
        IUIAnimationTransition *transition) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE AddKeyframeAtOffset( 
        /* [annotation][in] */ 
        UI_ANIMATION_KEYFRAME existingKeyframe,
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS offset,
        /* [annotation][retval][out] */ 
        UI_ANIMATION_KEYFRAME *keyframe) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE AddKeyframeAfterTransition( 
        /* [annotation][in] */ 
        IUIAnimationTransition *transition,
        /* [annotation][retval][out] */ 
        UI_ANIMATION_KEYFRAME *keyframe) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE AddTransitionAtKeyframe( 
        /* [annotation][in] */ 
        IUIAnimationVariable *variable,
        /* [annotation][in] */ 
        IUIAnimationTransition *transition,
        /* [annotation][in] */ 
        UI_ANIMATION_KEYFRAME startKeyframe) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE AddTransitionBetweenKeyframes( 
        /* [annotation][in] */ 
        IUIAnimationVariable *variable,
        /* [annotation][in] */ 
        IUIAnimationTransition *transition,
        /* [annotation][in] */ 
        UI_ANIMATION_KEYFRAME startKeyframe,
        /* [annotation][in] */ 
        UI_ANIMATION_KEYFRAME endKeyframe) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE RepeatBetweenKeyframes( 
        /* [annotation][in] */ 
        UI_ANIMATION_KEYFRAME startKeyframe,
        /* [annotation][in] */ 
        UI_ANIMATION_KEYFRAME endKeyframe,
        /* [annotation][in] */ 
        INT32 repetitionCount) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE HoldVariable( 
        /* [annotation][in] */ 
        IUIAnimationVariable *variable) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE SetLongestAcceptableDelay( 
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS delay) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE Schedule( 
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS timeNow,
        /* [annotation][defaultvalue][out] */ 
        UI_ANIMATION_SCHEDULING_RESULT *schedulingResult = 0) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE Conclude( void) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE Finish( 
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS completionDeadline) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE Abandon( void) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE SetTag( 
        /* [annotation][unique][in] */ 
        IUnknown *object,
        /* [annotation][in] */ 
        UINT32 id) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE GetTag( 
        /* [annotation][out] */ 
        IUnknown **object,
        /* [annotation][out] */ 
        UINT32 *id) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE GetStatus( 
        /* [annotation][retval][out] */ 
        UI_ANIMATION_STORYBOARD_STATUS *status) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE GetElapsedTime( 
        /* [annotation][out] */ 
        UI_ANIMATION_SECONDS *elapsedTime) = 0;
    
    virtual /* [annotation] */ 
    
    HRESULT STDMETHODCALLTYPE SetStoryboardEventHandler( 
        /* [annotation][unique][in] */ 
        IUIAnimationStoryboardEventHandler *handler) = 0;
    
};

MIDL_INTERFACE("DC6CE252-F731-41cf-B610-614B6CA049AD")
IUIAnimationTransition : public IUnknown
{
public:
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE SetInitialValue( 
        /* [annotation][in] */ 
        DOUBLE value) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE SetInitialVelocity( 
        /* [annotation][in] */ 
        DOUBLE velocity) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE IsDurationKnown( void) = 0;
    
    virtual /* [annotation] */ 
		
		HRESULT STDMETHODCALLTYPE GetDuration( 
        /* [annotation][retval][out] */ 
        UI_ANIMATION_SECONDS *duration) = 0;
        
};

MIDL_INTERFACE("274A7DEA-D771-4095-ABBD-8DF7ABD23CE3")
IUIAnimationTimerEventHandler : public IUnknown
{
public:
    virtual /* [annotation] */ 
		HRESULT STDMETHODCALLTYPE OnPreUpdate( void) = 0;
    
    virtual /* [annotation] */ 
		HRESULT STDMETHODCALLTYPE OnPostUpdate( void) = 0;
    
    virtual /* [annotation] */ 
		HRESULT STDMETHODCALLTYPE OnRenderingTooSlow( 
        /* [annotation][in] */ 
        UINT32 framesPerSecond) = 0;
};

MIDL_INTERFACE("195509B7-5D5E-4e3e-B278-EE3759B367AD")
IUIAnimationTimerUpdateHandler : public IUnknown
{
public:
    virtual /* [annotation] */ 
		HRESULT STDMETHODCALLTYPE OnUpdate( 
        /* [annotation][in] */ 
        UI_ANIMATION_SECONDS timeNow,
        /* [annotation][retval][out] */ 
        UI_ANIMATION_UPDATE_RESULT *result) = 0;
    
    virtual /* [annotation] */ 
		HRESULT STDMETHODCALLTYPE SetTimerClientEventHandler( 
        /* [annotation][in] */ 
        IUIAnimationTimerClientEventHandler *handler) = 0;
    
    virtual /* [annotation] */ 
		HRESULT STDMETHODCALLTYPE ClearTimerClientEventHandler( void) = 0;
        
};

#endif

class IBCGPAnimationTimerEventHandler : public IUIAnimationTimerEventHandler
{
protected:
	friend class CBCGPAnimationManager;
	
	IBCGPAnimationTimerEventHandler(CBCGPAnimationManager* pAnimManager)
	{
		ref = 1;
		m_pAnimationManager = pAnimManager;
	}
	HRESULT __stdcall OnPostUpdate()
	{
		return S_OK;
	}
	
	HRESULT __stdcall OnPreUpdate()
	{
		if (m_pAnimationManager != NULL)
		{
			m_pAnimationManager->OnWinAnimationTimerEvent();
		}
		
		return S_OK;
	}
	
	HRESULT __stdcall OnRenderingTooSlow(UINT32 /*framesPerSecond*/)
	{
		return E_NOTIMPL;
	}
	
	// IUnknown
	ULONG __stdcall AddRef()
	{
		return ++ref;
	}
	
	ULONG __stdcall Release()
	{
		return --ref;
	}
	HRESULT __stdcall QueryInterface(const IID& id,void**p)
	{
		if (id == IID_IUnknown || id == __uuidof(IUIAnimationTimerEventHandler))
		{
			*p = this;
			AddRef();
			return S_OK;
		}
		return E_NOINTERFACE;
	}
	
	CBCGPAnimationManager* m_pAnimationManager;
	unsigned long ref;
};

/////////////////////////////////////////////////////////////////////////////////
// CBCGPAnimationManager

template<class Interface>
AFX_INLINE void SafeRelease(Interface **ppInterfaceToRelease)
{
	if (*ppInterfaceToRelease != NULL)
	{
		(*ppInterfaceToRelease)->Release();
		(*ppInterfaceToRelease) = NULL;
	}
}

static const GUID BCGP_CLSID_UIAnimationManager		= { 0x4C1FC63A, 0x695C, 0x47E8, 0xA3, 0x39, 0x1A, 0x19, 0x4B, 0xE3, 0xD0, 0xB8 };
static const GUID BCGP_CLSID_UIAnimationTimer		= { 0xBFCD4A0C, 0x06B6, 0x4384, 0xB7, 0x68, 0x0D, 0xAA, 0x79, 0x2C, 0x38, 0x0E };
static const GUID BCGP_CLSID_UITransitionLibrary	= { 0x1D6322AD, 0xAA85, 0x4EF5, 0xA8, 0x28, 0x86, 0xD7, 0x10, 0x67, 0xD1, 0x45 };

#define WIN_ANIMATION_NOT_INITIALIZED	(-1)
#define WIN_ANIMATION_FAILED			(0)
#define WIN_ANIMATION_READY				(1)

CMap<UINT,UINT,CBCGPAnimationManager*,CBCGPAnimationManager*> CBCGPAnimationManager::m_mapManagers;
CCriticalSection CBCGPAnimationManager::g_cs;
int CBCGPAnimationManager::g_AnimationSupportedByOS = -1;

#define DEFAULT_TIMER_ELAPSE 75

BOOL CBCGPAnimationManager::InitWinAnimation()
{
	if (g_AnimationSupportedByOS == 0)
	{
		return FALSE;
	}

	if (m_nWinAnimationStatus != WIN_ANIMATION_NOT_INITIALIZED)
	{
		return m_nWinAnimationStatus == WIN_ANIMATION_READY ? TRUE : FALSE;
	}

	HRESULT hr = S_OK;

#ifndef _BCGSUITE_
	if (!globalData.ComInit())
	{
		return FALSE;
	}
#endif

	g_AnimationSupportedByOS = 0;
	m_nWinAnimationStatus = WIN_ANIMATION_FAILED;

	CoCreateInstance(
		BCGP_CLSID_UIAnimationManager,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_PPV_ARGS(&m_animationManager));
	if (m_animationManager == NULL)
	{
		CleanUp();
		return FALSE;
	}
	
	CoCreateInstance(
		BCGP_CLSID_UITransitionLibrary,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_PPV_ARGS(&m_transitionLibrary));
	if (m_transitionLibrary == NULL)
	{
		CleanUp();
		return FALSE;
	}

	CoCreateInstance(BCGP_CLSID_UIAnimationTimer, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_animationTimer));
	if (m_animationTimer == NULL)
	{
		CleanUp();
		return FALSE;
	}
	
	IUIAnimationTimerUpdateHandler* pAnimationTimerUpdateHandler = NULL;
	hr = m_animationManager->QueryInterface(__uuidof(IUIAnimationTimerUpdateHandler), (void**)&pAnimationTimerUpdateHandler);

	if (FAILED(hr))
	{
		CleanUp();
		return FALSE;
	}
		
	hr = m_animationTimer->SetTimerUpdateHandler(pAnimationTimerUpdateHandler, UI_ANIMATION_IDLE_BEHAVIOR_DISABLE);
	if (FAILED(hr))
	{
		CleanUp();
		return FALSE;
	}
	
	pAnimationTimerUpdateHandler->Release();
	pAnimationTimerUpdateHandler = NULL;
	
	m_pTimerEventHandler = new IBCGPAnimationTimerEventHandler(this);
	
	hr = m_animationTimer->SetTimerEventHandler(m_pTimerEventHandler);
	if (FAILED(hr))
	{
		CleanUp();
		return FALSE;
	}

	m_nWinAnimationStatus = WIN_ANIMATION_READY;
	g_AnimationSupportedByOS = 1;
	return TRUE;
}
//*****************************************************************************************
void CBCGPAnimationManager::CleanUp()
{
	SafeRelease(&m_variable);
	SafeRelease(&m_animationTimer);
	SafeRelease(&m_transitionLibrary);
	SafeRelease(&m_animationManager);
	
	if (m_pTimerEventHandler != NULL)
	{
		delete m_pTimerEventHandler;
		m_pTimerEventHandler = NULL;
	}
}
//*****************************************************************************************
CBCGPAnimationManager::CBCGPAnimationManager()
{
	m_bIsAnimated = FALSE;
	m_bInIdle = FALSE;
	m_bIsDelay = FALSE;
	m_animationManager = NULL;
	m_transitionLibrary = NULL;
	m_animationTimer = NULL;
	m_pTimerEventHandler = NULL;
	m_nWinAnimationStatus = WIN_ANIMATION_NOT_INITIALIZED;
	m_variable = NULL;
	m_dblAnimatedValue = 0.0;
	m_dblFinalValue = 0.0;
	m_nTimerID = 0;
	m_dblAnimationStep = 0.0;
	m_dblAnimationDuration = 0.0;
	m_AnimationType = BCGPANIMATION_Legacy;
	m_bIsFlashAnimation = FALSE;
	m_hrLastError = S_OK;
}
//*****************************************************************************************
CBCGPAnimationManager::~CBCGPAnimationManager()
{
	if (m_bIsAnimated)
	{
		StopAnimation(FALSE);
	}

	CleanUp();
}
//*****************************************************************************************
BOOL CBCGPAnimationManager::IsAnimationSupportedByOS()
{
	if (g_AnimationSupportedByOS == -1)
	{
		CBCGPAnimationManager animManager;
		animManager.InitWinAnimation();
	}
	return g_AnimationSupportedByOS == 1;
}
//*****************************************************************************************
BOOL CBCGPAnimationManager::StartAnimation(double dblStartValue, double dblFinalValue, double dblDurationInSeconds, BCGPAnimationType type, CBCGPAnimationManagerOptions* pOptions, double dblDelayInSeconds)
{
	StopAnimation(FALSE);

	m_hrLastError = S_OK;
	m_bIsDelay = FALSE;

	if (dblDurationInSeconds <= 0.0 || dblStartValue == dblFinalValue)
	{
		return FALSE;
	}

	m_bIsFlashAnimation = FALSE;
	m_dblAnimatedValue = dblStartValue;
	m_dblFinalValue = dblFinalValue;
	m_dblAnimationDuration = dblDurationInSeconds;
	m_AnimationType = type;

	if (type == BCGPANIMATION_Legacy || !InitWinAnimation())
	{
		m_bIsAnimated = TRUE;

		m_bIsDelay = dblDelayInSeconds > 0;

		const int nElapse = m_bIsDelay ? (int)(dblDelayInSeconds * 1000) : DEFAULT_TIMER_ELAPSE;
		const double dblCount = m_dblAnimationDuration * 1000 / DEFAULT_TIMER_ELAPSE - 1.0;

		if (dblCount > 0.)
		{
			m_dblAnimationStep = (m_dblFinalValue - dblStartValue) / dblCount;
		}
		else
		{
			m_dblAnimationStep = (m_dblFinalValue - dblStartValue);
		}

		m_nTimerID = (UINT)::SetTimer(NULL, 0, nElapse, AnimationTimerProc);
		
		g_cs.Lock ();
		m_mapManagers.SetAt(m_nTimerID, this);
		g_cs.Unlock ();

		return TRUE;
	}

	CBCGPAnimationManagerOptions options;
	if (pOptions != NULL)
	{
		options = *pOptions;
	}

	if (m_variable == NULL)
	{
		m_hrLastError = m_animationManager->CreateAnimationVariable(dblStartValue, &m_variable);
		if (FAILED(m_hrLastError))
		{
			return FALSE;
		}
	}

	IUIAnimationStoryboard* storyboard = NULL;
	IUIAnimationTransition* transition = NULL;
	IUIAnimationTransition* transitionDelay = NULL;

	m_hrLastError = m_animationManager->CreateStoryboard(&storyboard);
	if (SUCCEEDED(m_hrLastError))
	{
		switch (type)
		{
		case BCGPANIMATION_AccelerateDecelerate:
			m_hrLastError = m_transitionLibrary->CreateAccelerateDecelerateTransition(
				dblDurationInSeconds, dblFinalValue, 
				options.m_dblAccelerationRatio, options.m_dblDecelerationRatio, &transition);
			break;

		case BCGPANIMATION_Cubic:
		case BCGPANIMATION_ParabolicFromAcceleration:
			{
				double dblRatio = (type == BCGPANIMATION_Cubic) ? 1.44 : 0.7;

				double dblFinalVelocity = options.m_dblFinalVelocity == 0.0 ? 
					dblRatio * fabs(dblFinalValue - dblStartValue) / dblDurationInSeconds : options.m_dblFinalVelocity;

				if (dblFinalValue > dblStartValue)
				{
					dblFinalVelocity = -dblFinalVelocity;
				}

				if (type == BCGPANIMATION_Cubic)
				{
					m_hrLastError = m_transitionLibrary->CreateCubicTransition(
						dblDurationInSeconds, dblFinalValue, dblFinalVelocity, &transition);
				}
				else
				{
					double dblParabolicAccelerationRatio = options.m_dblParabolicAccelerationRatio;
					if (dblParabolicAccelerationRatio == 0.0)
					{
						dblParabolicAccelerationRatio = 10.0 * fabs(dblFinalVelocity / dblDurationInSeconds);
					}

					m_hrLastError = m_transitionLibrary->CreateParabolicTransitionFromAcceleration(
						dblFinalValue, dblFinalVelocity, dblParabolicAccelerationRatio, &transition);
				}
			}
			break;

		case BCGPANIMATION_Linear:
			m_hrLastError = m_transitionLibrary->CreateLinearTransition(
				dblDurationInSeconds, dblFinalValue, &transition);
			break;

		case BCGPANIMATION_SmoothStop:
			m_hrLastError = m_transitionLibrary->CreateSmoothStopTransition(
				dblDurationInSeconds, dblFinalValue, &transition);
			break;
		}

		if (SUCCEEDED(m_hrLastError))
		{
			m_animationTimer->Enable();

			m_hrLastError = storyboard->AddTransition(m_variable, transition);

			UI_ANIMATION_SECONDS secondsNow = static_cast<UI_ANIMATION_SECONDS>(0);
			if (SUCCEEDED(m_hrLastError))
			{
				m_hrLastError = m_animationTimer->GetTime(&secondsNow);
				if (SUCCEEDED(m_hrLastError))
				{
					m_hrLastError = storyboard->Schedule(secondsNow + dblDelayInSeconds);
				}
			}
		}
	}

	SafeRelease(&transition);
	SafeRelease(&transitionDelay);
	SafeRelease(&storyboard);

	if (SUCCEEDED(m_hrLastError))
	{
		m_bIsAnimated = TRUE;
		OnAnimationValueChanged(dblStartValue, dblStartValue);
		return TRUE;
	}

	TRACE(_T("CBCGPAnimationManager::Start failed. Error code: %x\n"), m_hrLastError);
	return FALSE;
}
//*****************************************************************************************
BOOL CBCGPAnimationManager::StartFlashAnimation(double dblPeriodInSeconds)
{
	StopAnimation(FALSE);

	if (dblPeriodInSeconds <= 0.0)
	{
		return FALSE;
	}

	m_AnimationType = BCGPANIMATION_Legacy;
	m_dblAnimatedValue = 0.0;
	m_dblFinalValue = 0.0;
	m_dblAnimationDuration = 0.0;
	m_bIsAnimated = TRUE;
	m_dblAnimationStep = 0.0;
	m_bIsFlashAnimation = TRUE;
	m_nTimerID = (UINT)::SetTimer(NULL, 0, (int)(1000 * dblPeriodInSeconds), AnimationTimerProc);
	
	g_cs.Lock ();
	m_mapManagers.SetAt(m_nTimerID, this);
	g_cs.Unlock ();

	return TRUE;
}
//*****************************************************************************************
void CBCGPAnimationManager::OnWinAnimationTimerEvent()
{
	if (m_variable == NULL)
	{
		return;
	}

	UI_ANIMATION_SECONDS secondsNow = 0;
	HRESULT hr = m_animationTimer->GetTime(&secondsNow);

	if (SUCCEEDED(hr))
	{
		hr = m_animationManager->Update(secondsNow);

		if (SUCCEEDED(hr))
		{
			double value;
			m_variable->GetValue(&value);

			if (value != m_dblAnimatedValue)
			{
				OnAnimationValueChanged(m_dblAnimatedValue, value);
				m_dblAnimatedValue = value;
			}
		}
	}

	UI_ANIMATION_MANAGER_STATUS status;
	hr = m_animationManager->GetStatus(&status);

	if (SUCCEEDED(hr))
	{
		if (status == UI_ANIMATION_MANAGER_IDLE)
		{
			m_bIsAnimated = FALSE;
			OnAnimationFinished();
		}
		else if (!m_bInIdle)
		{
			m_bInIdle = TRUE;
			OnAnimationIdle();
			m_bInIdle = FALSE;
		}
	}
}
//*****************************************************************************************
void CBCGPAnimationManager::OnTimerEvent()
{
	if (m_nTimerID == 0)
	{
		return;
	}

	if (!m_bInIdle && m_bIsAnimated)
	{
		m_bInIdle = TRUE;
		OnAnimationIdle();
		m_bInIdle = FALSE;
	}

	double dblOldValue = m_dblAnimatedValue;

	if (m_bIsFlashAnimation)
	{
		m_dblAnimatedValue = dblOldValue == 0.0 ? 1.0 : 0.0;
		OnAnimationValueChanged(dblOldValue, m_dblAnimatedValue);

		return;
	}

	if (m_bIsDelay)
	{
		m_bIsDelay = FALSE;

		::KillTimer(NULL, m_nTimerID);
		
		g_cs.Lock ();
		
		m_mapManagers.RemoveKey(m_nTimerID);

		m_nTimerID = (UINT)::SetTimer(NULL, 0, DEFAULT_TIMER_ELAPSE, AnimationTimerProc);
		m_mapManagers.SetAt(m_nTimerID, this);

		g_cs.Unlock ();

		return;
	}

	m_dblAnimatedValue += m_dblAnimationStep;
	
	BOOL bStopTransition = FALSE;
	
	if (m_dblAnimationStep < 0)
	{
		if (m_dblAnimatedValue <= m_dblFinalValue)
		{
			m_dblAnimatedValue = m_dblFinalValue;
			bStopTransition = TRUE;
		}
	}
	else
	{
		if (m_dblAnimatedValue >= m_dblFinalValue)
		{
			m_dblAnimatedValue = m_dblFinalValue;
			bStopTransition = TRUE;
		}
	}
	
	OnAnimationValueChanged(dblOldValue, m_dblAnimatedValue);
	
	if (bStopTransition)
	{
		StopAnimation(TRUE);
	}
}
//*****************************************************************************************
void CBCGPAnimationManager::StopAnimation(BOOL bNotify/* = TRUE*/)
{
	if (m_nTimerID != 0)
	{
		::KillTimer(NULL, m_nTimerID);
		
		g_cs.Lock ();
		m_mapManagers.RemoveKey(m_nTimerID);
		g_cs.Unlock ();
		
		m_nTimerID = 0;
	}

	if (m_nWinAnimationStatus == WIN_ANIMATION_READY && m_animationManager != NULL)
	{
		UI_ANIMATION_MANAGER_STATUS status;
		HRESULT hr = m_animationManager->GetStatus(&status);
		
		if (SUCCEEDED(hr))
		{
			m_animationManager->FinishAllStoryboards(1);
		}

		SafeRelease(&m_variable);
		m_animationTimer->Disable();
	}

	m_bIsAnimated = FALSE;

	if (bNotify)
	{
		OnAnimationFinished();
	}
}
//*****************************************************************************************
VOID CALLBACK CBCGPAnimationManager::AnimationTimerProc(HWND /*hwnd*/, UINT /*uMsg*/, UINT_PTR idEvent, DWORD /*dwTime*/)
{
	CBCGPAnimationManager* pObject = NULL;
	
	g_cs.Lock();
	
	if (!m_mapManagers.Lookup((UINT)idEvent, pObject))
	{
		g_cs.Unlock ();
		return;
	}
	
	ASSERT(pObject != NULL);
	
	g_cs.Unlock();
	
	pObject->OnTimerEvent();
}
