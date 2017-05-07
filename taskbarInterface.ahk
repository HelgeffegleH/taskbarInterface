class taskbarInterface {
	__new(hwnd,onButtonClickFunction){
		this.hwnd:=hwnd										; Handle to the window whose taskbar preview will recieve the buttons
		this.callback:= IsObject(onButtonClickFunction) ? onButtonClickFunction : Func(onButtonClickFunction)	; Pass a func object or function name.
		taskbarInterface.allInterfaces[hwnd]:=this			; allInterfaces array is used for routing the callbacks.
		if !taskbarInterface.init							; On first call here, initialise the com object and turn on button messages. (WM_COMMAND)
			taskbarInterface.initCom(), taskbarInterface.turnOnButtonMessages()
		this.createButtons()								; Create the buttons for this interface.
		this.isDisabled:=false								; By deafault, the interface is not disabled, us stopThisButtonMonitor() to disable.
	}
	; Context: this = "new taskbarInterface(...)"
	; Note, further down the context switches to this=taskbarInterface, for convenience. The switch is clearly marked.
	; User methods.
	; in the following n is the button number, n ∈ [1,7]⊂ℤ
	showButton(n){
		; Show button n
		static THBF_HIDDEN:=0x8
		this.updateThumbButtonFlags(n,0,THBF_HIDDEN)		; Add flag 0, remove 0x8 (THBF_HIDDEN)
		return this.ThumbBarUpdateButtons(n)				; Update
	}
	hideButton(n){
		; Hide button n
		static THBF_HIDDEN:=0x8
		this.updateThumbButtonFlags(n,THBF_HIDDEN,0)		; Update flag THBF_HIDDEN:=0x8
		return this.ThumbBarUpdateButtons(n)				; Update
	}
	disableButton(n){
		; disable button n
		; The button becomes unclickable and grays out.
		static THBF_DISABLED:=0x1
		this.updateThumbButtonFlags(n,THBF_DISABLED,0)		; Update flag, add THBF_DISABLED:=0x8
		return this.ThumbBarUpdateButtons(n)				; Update
	}
	enableButton(n){
		; reenable button n
		; the button becomes clickable and regains its
		; color.
		static THBF_DISABLED:=0x1
		this.updateThumbButtonFlags(n,0,THBF_DISABLED)		; Update flag, remove THBF_DISABLED:=0x8
		return this.ThumbBarUpdateButtons(n)				; Update
	}
	setButtonIcon(n,hIcon){
		; Set button icon for button n
		; Call queryButtonIconSize() to obtain the required size of the icon. See queryButtonIconSize() below.
		static THB_ICON:=0x2					
		this.updateThumbButtonMask(n,THB_ICON,0)			; Update mask THB_ICON
		this.setThumbButtonhIcon(n,hIcon)					; Set the icon handle
		return this.ThumbBarUpdateButtons(n)				; Update
	}
	setButtonToolTip(n,text:=""){
		; Sets the tooltip for button n, that is shown when the
		; mouse cursors hover over the button for a few seconds.
		static THB_TOOLTIP:=0x4
		if (text!="")
			this.updateThumbButtonMask(n,THB_TOOLTIP,0)		; Update mask THB_TOOLTIP, add
		this.setThumbButtonToolTipText(n,text)				; Update the text
		this.ThumbBarUpdateButtons(n)						; Update
		if (text="")
			this.updateThumbButtonMask(n,0,THB_TOOLTIP)		; Update mask THB_TOOLTIP, remove
		return 
	}
	dismissPreviewOnButtonClick(n,dismiss:=true){	
        ; Call with dismiss:=true to make button  n's  click
        ; to  cause  the thumbnail preview to be dismissed
        ; (close). To show again, hover mouse on taskbar icon.
        ; Call with  dismiss:false  to  invoke  the  default
		; behaviour, i.e, no dismiss on click.
		static THBF_DISMISSONCLICK  := 0x2
		if dismiss
			this.updateThumbButtonFlags(n,THBF_DISMISSONCLICK,0)		; Update flag, add THBF_DISMISSONCLICK:=0x2
		else
			this.updateThumbButtonFlags(n,0,THBF_DISMISSONCLICK)		; Update flag, remove THBF_DISMISSONCLICK:=0x2
		return this.ThumbBarUpdateButtons(n)							; Update
	}
	removeButtonBackground(n){
		; Remove the background (and/or border) of button n.
		; The button has a background by default
		static THBF_NOBACKGROUND  := 0x4
		this.updateThumbButtonFlags(n,THBF_NOBACKGROUND,0)				; Update flag, add THBF_NOBACKGROUND:=0x2
		return this.ThumbBarUpdateButtons(n)							; Update
	}
	reAddButtonBackground(n){
		; Readd the background (and/or) border of button n.
		; Only needed to call if removeButtonBackground(n) was
		; previously called
		static THBF_NOBACKGROUND  := 0x4
		this.updateThumbButtonFlags(n,0,THBF_NOBACKGROUND)				; Update flag, remove THBF_NOBACKGROUND:=0x2
		return this.ThumbBarUpdateButtons(n)							; Update
	}
	
	setButtonNonInteractive(n){
		; Set button n to be non-interactive,
		; similar to disableButton(n), but the button doesn't gray out.
		static THBF_NONINTERACTIVE  := 0x10
		this.updateThumbButtonFlags(n,THBF_NONINTERACTIVE,0)			; Update flag, add THBF_NONINTERACTIVE:=0x10
		return this.ThumbBarUpdateButtons(n)							; Update
	}
	setButtonInteractive(n){
		; Set button n to be interactive again.
		static THBF_NONINTERACTIVE  := 0x10
		this.updateThumbButtonFlags(n,0,THBF_NONINTERACTIVE)			; Update flag, remove THBF_NONINTERACTIVE:=0x10
		return this.ThumbBarUpdateButtons(n)							; Update
	}
	; Misc
	freeMemory(){
		; There is no need to keep the memory allocated if there is no intentions to make any changes to the interface.
		if this.THUMBBUTTON
			this.GlobalFree(this.THUMBBUTTON)
		this.isFreed:=true
		return
	}
	clear(){
		; Remove from allInterfaces array
		taskbarInterface.allInterfaces.Delete(this.hwnd)
		; Free memory
		this.freeMemory()
		return
	}
	stopThisButtonMonitor(){
		; This will dismiss the callback, the message monitor is still on. To turn off message monitor use stopAllButtonMonitor()
		; Default is message monitoring on
		return this.isDisabled:=true
	}
	restartThisButtonMonitor(){
		; This will reenable the button click callbacks. If all message monitor is off, i.e., stopAllButtonMonitor() was called, restart by calling restartAllButtonMonitor()
		; Default is message monitoring on
		return this.isDisabled:=false
	}
	; Global methods, affects all interfaces
	clearAll(){
		for k, interface in taskbarInterface.allInterfaces
			interface.clear()
		return
	}
	stopAllButtonMonitor(){
		; turns off button message monitor, default is on.
		if taskbarInterface.allDisabled
			return
		return taskbarInterface.turnOffButtonMessages(), taskbarInterface.allDisabled:=true
	}
	restartAllButtonMonitor(){
		; turns on button message monitor, deafult is on, hence no need to call if you didn't turn it off
		if taskbarInterface.allDisabled
			return taskbarInterface.turnOnButtonMessages(), taskbarInterface.allDisabled:=false
	}
	queryButtonIconSize(){	
		; Returns the required pixel width and height for the button icons.
		; Example:
		; sz:=taskbarInterface.queryButtonIconSize()
		; Msgbox, % "The icon width must be: " sz.w  "`nThe icon height must be: " sz.h
		SysGet, SM_CXICON, 11
		SysGet, SM_CYICON, 12
		return {w:SM_CXICON, h:SM_CYICON}
	}
	; End user methods
	; Internal methods
	; Meta functions 
	__Call(fn,p*){
		; For verifying correct input. maybe change this.
		if InStr( 	  ",showButton,hideButton,setButtonIcon,enableButton,disableButton"
					. ",setButtonToolTip,dismissPreviewOnButtonClick,removeButtonBackground"
					. ",reAddButtonBackground,setButtonNonInteractive,setButtonInteractive,", "," . fn . ",") {
			if this.isFreed
				throw Exception("This interface has freed its memory, it cannot be used.",-1) 					; If the user tries to alter the apperance or function of the interface after memory was free, throw an exception.
			this.verifyiId(p[1])
		}
	}
	__Delete(){
		msgbox, hi debug
		return
	}
	verifyiId(iId){
		; Ensures the button number iId, is in the correct range.
		; Avoids unexpected behaviour by passing an address outside of allocated memory in this.THUMBBUTTON
		if (iId<1 || iId>7 || round(iId)!=iId)
			throw Exception("Button number must be an integer in the in range 1 to 7 (inclusive)",-2)
		return 1
	}
	createButtons(){
		; Creates 7 buttons. All hidden. This is because ThumbBarAddButtons() can only be called once, it seems. This is for convenience.
		; All buttons will have the THB_FLAGS mask.
		static THB_FLAGS:=0x00000008
		static THBF_HIDDEN:=0x8
		this.THUMBBUTTON:=this.GlobalAlloc(this.thumbButtonSize*7)
		
		loop, 7 {
			structOffset:=this.thumbButtonSize*(A_Index-1)
			NumPut(A_Index,this.THUMBBUTTON+structOffset, 4, "Uint")					; Specify the ids: 1,...,7
			this.updateThumbButtonMask(A_Index,THB_FLAGS,0)								; update the mask: THB_FLAGS
			this.updateThumbButtonFlags(A_Index,THBF_HIDDEN,0)							; Update flag: THBF_HIDDEN:=0x8
		}
		return this.ThumbBarAddButtons()
	}

	; Update/get/set methods for the THUMBBUTTON struct array.
	; The update functions call the get functions, modifies the values and then set. The caller of update() then calls ThumbBarUpdateButtons() when finished
	
	; Update
	updateThumbButtonMask(iId,add:=0,remove:=0){
		dwMask:=(this.getThumbButtonMask(iId)|add)^remove
		return this.setThumbButtonMask(iId,dwMask)
	}
	updateThumbButtonFlags(iId,add:=0,remove:=0){
		dwflags:=(this.getThumbButtonFlags(iId)|add)^remove
		return this.setThumbButtonFlags(iId,dwFlags)
	}
	; Item and struct offsets are specified for maintainabillity
	; Write values to adress at this.THUMBBUTTON + structOffset + itemOffset
	; Get
	getThumbButtonMask(iId){
		static	itemOffset		:=	0																		; dwMask
				structOffset	:=	this.thumbButtonSize*(iId-1)
		return NumGet(this.THUMBBUTTON+itemOffset+structOffset, "Uint")
	}
	getThumbButtonFlags(iId){
		static	itemOffset		:=	8+2*A_PtrSize+260*2														; dwFlags
				structOffset	:=	this.thumbButtonSize*(iId-1)	
		return NumGet(this.THUMBBUTTON+itemOffset+structOffset,0,"Uint")
	}
	; Set
	setThumbButtonMask(iId,dwMask){
		static	itemOffset		:=	0																		; dwMask
				structOffset	:=	this.thumbButtonSize*(iId-1)
		return NumPut(dwMask, this.THUMBBUTTON+itemOffset+structOffset, "Uint")
	}
	setThumbButtonhIcon(iId,hIcon){
		static	itemOffset		:=	8+A_PtrSize																; hIcon
				structOffset	:=	this.thumbButtonSize*(iId-1)
		return NumPut(hIcon, this.THUMBBUTTON+itemOffset+structOffset, "Ptr")
	}
	setThumbButtonToolTipText(iId,text:=""){
		static	itemOffset		:=	8+2*A_PtrSize															; szTip
				structOffset	:=	this.thumbButtonSize*(iId-1)
		;if (text="")
			;return NumPut(0, this.THUMBBUTTON+structOffset+itemOffset, 0, "Char")							; Null terminate, there is room for an int.
		return StrPut(SubStr(text,1,259), this.THUMBBUTTON+structOffset+itemOffset, 260, "UTF-16")			; Make sure tooltip text isn't too long
	}
	setThumbButtonFlags(iId,dwFlags){
		static	itemOffset		:=	8+2*A_PtrSize+260*2														; dwFlags
				structOffset	:=	this.thumbButtonSize*(iId-1)	
		return NumPut(dwFlags, this.THUMBBUTTON+structOffset+itemOffset, "Uint")
	}
	
	;
	; Com Interface wrapper functions
	; The bound funcs are made in init()
	ThumbBarAddButtons() {
		; This function can only be called once it seems. Make one call and add all buttons hidden. Then use ThumbBarUpdateButtons() to "add" and "remove" buttons via the THBF_HIDDEN flag
		; Max buttons is 7
		return taskbarInterface.ThumbBarAddButtonsFn.Call("Ptr", this.hWnd, "Uint", 7, "Ptr", this.THUMBBUTTON) ; return 0 is ok!
	}
	
	ThumbBarUpdateButtons(iId){
		return taskbarInterface.ThumbBarUpdateButtonsFn.Call("Ptr", this.hWnd, "Uint", 1, "Ptr", this.THUMBBUTTON+this.thumbButtonSize*(iId-1)) ; return 0 is ok!
	}
	ThumbBarToolTip(){
		return taskbarInterface.ThumbBarToolTipFn.Call("Ptr", this.hWnd, "Str", this.tooltipText)
	}
	             
	;
	; Static variables
	;
	static allInterfaces:=[] 						; Tracks all interfaces, for callbacks.
	static init:=0									; For first time use initialising of the com object.
	
	; THUMBBUTTON  struct:
	static thumbButtonSize:=A_PtrSize=4?544:552		; Size calculations according to:
	/*
	; URL:
	;	- https://msdn.microsoft.com/en-us/library/windows/desktop/dd391559(v=vs.85).aspx (THUMBBUTTON structure)
	;
	;									offsets:							Contribution to size (bytes):	
	THUMBBUTTONMASK  dwMask				0						...			4
	UINT             iId				4						...			4
	UINT             iBitmap			8						...			4
	;															...			64-bit: add 4 bytes spacing, pointer address, adr, needs to be mod(adr,A_PtrSize)=0
	HICON            hIcon				8+A_PtrSize				...			A_PtrSize
	WCHAR            szTip[260]			8+2*A_PtrSize			...			260*2
	THUMBBUTTONFLAGS dwFlags			8+2*A_PtrSize+260*2		...			4
	;																		Sum: 32-bit: 4+4+4+0+A_PtrSize+260*2+4=544, 544/A_PtrSize=136, no spacing needed.
	;																		Sum: 64-bit: 4+4+4+4+A_PtrSize+260*2+4=548, 548/A_PtrSize=68.5 -> add 4 bytes, 552/A_PtrSize=69 (mod(552,A_Ptrsize)=0).
	;																		64-bit: add 4 bytes spacing to next struct in array
	;																		Conclusion: size:= A_PtrSize=4?544:552
	
	*/
	;
	; NOTE:
	;
	; 			<	>	<	>	<	>	<	>	<	>	<	>					Context:						<	>	<	>	<	>	<	>	<	>	<	>
	; 			<	>	<	>	<	>	<	>	<	>	<	>			 this = taskbarInterface				<	>	<	>	<	>	<	>	<	>	<	>
	; 			<	>	<	>	<	>	<	>	<	>	<	>													<	>	<	>	<	>	<	>	<	>	<	>
	;
	;
	initCom(){
		; Url:
		;	-  https://msdn.microsoft.com/en-us/library/windows/desktop/bb774652(v=vs.85).aspx (ITaskbarList interface)
		; Initilises the com object.
		static CLSID_TaskbarList := "{56FDF344-FD6D-11d0-958A-006097C9A090}"
		static IID_ITaskbarList3 := "{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}"
		this.hComObj := ComObjCreate(CLSID_TaskbarList, IID_ITaskbarList3)
		; Get the address to the vTable.
		this.vTable:=NumGet(this.hComObj+0)
		; Create function objects for the interface, for convenience and clarity
																																								; Name:					 Number:
		this.ThumbBarAddButtonsFn:=Func("DllCall").Bind(NumGet(this.vTable+15*A_PtrSize,0,"Ptr"), "Ptr", this.hComObj)											; ThumbBarAddButtons		(15)
		this.ThumbBarUpdateButtonsFn:=Func("DllCall").Bind(NumGet(this.vTable+16*A_PtrSize,0,"Ptr"), "Ptr", this.hComObj)										; ThumbBarUpdateButtons		(16)
		this.ThumbBarToolTipFn:=Func("DllCall").Bind(NumGet(this.vTable+19*A_PtrSize,0,"Ptr"), "Ptr", this.hComObj)												; ThumbBarToolTip			(19)
																																		
		this.init:=1
		return	
	}
	
	
	; Click on button handling:
	; URL:
	;	- https://msdn.microsoft.com/en-us/library/windows/desktop/dd391703(v=vs.85).aspx (ITaskbarList3::ThumbBarAddButtons method, remarks)
	;		When a button in a thumbnail toolbar is clicked, the window associated with that thumbnail is sent a WM_COMMAND
	;		message with the HIWORD of its wParam parameter set to THBN_CLICKED and the LOWORD to the button ID.
	;
	turnOffButtonMessages(){
		static WM_COMMAND := 0x111
		if this.buttonMessageFn
			OnMessage(WM_COMMAND,this.buttonMessageFn,0)					; Turn off button message monitoring. 
		return
	}
	turnOnButtonMessages(){
		static WM_COMMAND := 0x111
		if !this.buttonMessageFn
			this.buttonMessageFn:=ObjBindMethod(this,"onButtonClick")	; The monitor function is kept for 
		OnMessage(WM_COMMAND,this.buttonMessageFn) 						; When the buttons are clicked, a WM_COMMAND message is sent.
		return
	}
	onButtonClick(wParam,lParam,msg,hWnd){
		; HIWORD of wParam is THBN_CLICKED when a button was clicked.  	(wParam>>16)
		; LOWORD of wParam is the button number.						(wParam&0xffff)
		static THBN_CLICKED := 0x1800
		Critical, On
		if (wParam >> 16 = THBN_CLICKED) && taskbarInterface.allInterfaces.HasKey(hWnd) {
			ref:=taskbarInterface.allInterfaces[hWnd] 				;	The reference to the interface whose button was clicked
			if ref.isDisabled
				return 1
			buttonNumber:= wParam&0xffff
			tf:=ref.callback.Bind(buttonNumber,ref) 				;	The callback includes the button number and a reference to the interface.
			SetTimer, % tf, -1										;	The call is delayed until this function has returned.
			return 1												;	No further handling of this message is needed, asumption.
		}
		return
	}
	;	NOTE: End context: this=taskbarInterface
	; 		
	;
	; Memory allocation methods.
	GlobalAlloc(dwBytes){
		; URL:
		;	- https://msdn.microsoft.com/en-us/library/windows/desktop/aa366574(v=vs.85).aspx (GlobalAlloc function)
		static GMEM_ZEROINIT:=0x0040	; Zero fill memory
		static uFlags:=GMEM_ZEROINIT	; For clarity.
		h:=DllCall("Kernel32.dll\GlobalAlloc", "Uint", uFlags, "Ptr", dwBytes, "Ptr")
		if !h
			throw Exception("Memory alloc failed.",-1)
		return h
	}
	GlobalFree(hMem){
		; URL:
		;	- https://msdn.microsoft.com/en-us/library/windows/desktop/aa366579(v=vs.85).aspx (GlobalFree function)
		h:=DllCall("Kernel32.dll\GlobalFree", "Ptr", hMem, "Ptr")
		if h
			throw Exception("Memory free failed",-1)
		return h
	}
	; Additional reference:
	; ShObjIdl.h:
	/*
	 typedef struct ITaskbarList3Vtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 									(0)
            __RPC__in ITaskbarList3 * This,
             [in]  __RPC__in REFIID riid,
             [annotation][iid_is][out]  
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 											(1)
            __RPC__in ITaskbarList3 * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )(											(2) 
            __RPC__in ITaskbarList3 * This);
        
        HRESULT ( STDMETHODCALLTYPE *HrInit )( 											(3)
            __RPC__in ITaskbarList3 * This);
        
        HRESULT ( STDMETHODCALLTYPE *AddTab )( 											(4)
            __RPC__in ITaskbarList3 * This,
             [in]  __RPC__in HWND hwnd);
        
        HRESULT ( STDMETHODCALLTYPE *DeleteTab )( 										(5)
            __RPC__in ITaskbarList3 * This,
             [in]  __RPC__in HWND hwnd);
        
        HRESULT ( STDMETHODCALLTYPE *ActivateTab )( 									(6)
            __RPC__in ITaskbarList3 * This,
             [in]  __RPC__in HWND hwnd);
        
        HRESULT ( STDMETHODCALLTYPE *SetActiveAlt )(									(7)
            __RPC__in ITaskbarList3 * This,
             [in]  __RPC__in HWND hwnd);
        
        HRESULT ( STDMETHODCALLTYPE *MarkFullscreenWindow )( 							(8)
            __RPC__in ITaskbarList3 * This,
             [in]  __RPC__in HWND hwnd,
             [in]  BOOL fFullscreen);
        
        HRESULT ( STDMETHODCALLTYPE *SetProgressValue )( 								(9)
            __RPC__in ITaskbarList3 * This,
             [in]  __RPC__in HWND hwnd,
             [in]  ULONGLONG ullCompleted,
             [in]  ULONGLONG ullTotal);
        
        HRESULT ( STDMETHODCALLTYPE *SetProgressState )( 								(10)
            __RPC__in ITaskbarList3 * This,
             [in]  __RPC__in HWND hwnd,
             [in]  TBPFLAG tbpFlags);
        
        HRESULT ( STDMETHODCALLTYPE *RegisterTab )( 									(11)
            __RPC__in ITaskbarList3 * This,
             [in]  __RPC__in HWND hwndTab,
             [in]  __RPC__in HWND hwndMDI);
        
        HRESULT ( STDMETHODCALLTYPE *UnregisterTab )( 									(12)
            __RPC__in ITaskbarList3 * This,
             [in]  __RPC__in HWND hwndTab);
        
        HRESULT ( STDMETHODCALLTYPE *SetTabOrder )( 									(13)
            __RPC__in ITaskbarList3 * This,
             [in]  __RPC__in HWND hwndTab,
             [in]  __RPC__in HWND hwndInsertBefore);
        
        HRESULT ( STDMETHODCALLTYPE *SetTabActive )( 									(14)
            __RPC__in ITaskbarList3 * This,
             [in]  __RPC__in HWND hwndTab,
             [in]  __RPC__in HWND hwndMDI,
             [in]  DWORD dwReserved);
        
        HRESULT ( STDMETHODCALLTYPE *ThumbBarAddButtons )( 								(15)
            __RPC__in ITaskbarList3 * This,
             [in]  __RPC__in HWND hwnd,
             [in]  UINT cButtons,
             [size_is][in]  __RPC__in_ecount_full(cButtons) LPTHUMBBUTTON pButton);    
        
        HRESULT ( STDMETHODCALLTYPE *ThumbBarUpdateButtons )( 							(16)
            __RPC__in ITaskbarList3 * This,
             [in]  __RPC__in HWND hwnd,
             [in]  UINT cButtons,
             [size_is][in]  __RPC__in_ecount_full(cButtons) LPTHUMBBUTTON pButton);
        
        HRESULT ( STDMETHODCALLTYPE *ThumbBarSetImageList )( 							(17)
            __RPC__in ITaskbarList3 * This,
             [in]  __RPC__in HWND hwnd,
             [in]  __RPC__in_opt HIMAGELIST himl);
        
        HRESULT ( STDMETHODCALLTYPE *SetOverlayIcon )( 									(18)
            __RPC__in ITaskbarList3 * This,
             [in]  __RPC__in HWND hwnd,
             [in]  __RPC__in HICON hIcon,
             [string][unique][in]  __RPC__in_opt_string LPCWSTR pszDescription);
        
        HRESULT ( STDMETHODCALLTYPE *SetThumbnailTooltip )( 							(19)
            __RPC__in ITaskbarList3 * This,
             [in]  __RPC__in HWND hwnd,
             [string][unique][in]  __RPC__in_opt_string LPCWSTR pszTip);
        
        HRESULT ( STDMETHODCALLTYPE *SetThumbnailClip )( 								(20)
            __RPC__in ITaskbarList3 * This,
             [in]  __RPC__in HWND hwnd,
             [in]  __RPC__in RECT *prcClip);
        
        END_INTERFACE
    } ITaskbarList3Vtbl;
	*/
}

