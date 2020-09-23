VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'        _/_/_/      _/_/_/  _/      _/        _/_/_/  _/_/_/  _/_/_/
'       _/    _/  _/        _/_/  _/_/      _/          _/      _/
'      _/_/_/    _/        _/  _/  _/      _/  _/_/    _/      _/
'     _/    _/  _/        _/      _/      _/    _/    _/      _/
'    _/    _/    _/_/_/  _/      _/        _/_/_/  _/_/_/  _/_/_/


'****************************************************************************************
'*  RCM GII!     Radical Client and Non Client Modification Class - Gen II v2.7.4       *
'*                                                                                      *
'*  Created:     August 1, 2006, (original v1. Febuary 2006)                            *
'*  Updated:     September 16, 2006                                                     *
'*  Purpose:     Form and Control Modification Class -with WinSubHookGX.tlb             *
'*  Functions:   (listed)                                                               *
'*  Revision:    2.7.4                                                                  *
'*  Compile:     Native                                                                 *
'*  Author:      John Underhill (Steppenwolfe)                                          *
'*  Credit:      Based in part on Neocaption written by Steve McMahon                   *
'*                                                                                      *
'****************************************************************************************


'*~ List of exposed functions ~*
'/~ Attatch                     - start
'/~ Clean_Up                    - manually destroy resources (not used)
'/  Change_Caption              - change the forms caption

'/~ Form Properties /~
'/~ p_ICaption                  - caption bar image                     (img)
'/~ p_ICBottom                  - bottom border image                   (img)
'/~ p_ICLeft                    - left border image                     (img)
'/~ p_ICRight                   - right border image                    (img)
'/~ p_ICBoxMin                  - control box minimize                  (img)
'/~ p_ICBoxMax                  - control box maximize                  (img)
'/~ p_ICBoxRst                  - control box restore                   (img)
'/~ p_ICBoxCls                  - control box close                     (img)
'/~ p_BottomSizingBorder        - bottom sizing border                  (long)
'/~ p_ButtonOffsetX             - control button offset horz            (long)
'/~ p_ButtonOffsetY             - control button offset vert            (long)
'/~ p_CustomCaption             - use transparent regions               (bool)
'/~ p_CaptionFntClr             - caption text font color               (long)
'/~ p_CaptionFrame              - use a caption bar text frame image    (bool)
'/~ p_CaptionOffset             - caption title offset                  (long)
'/~ p_CenterCaption             - center caption on form                (bool)
'/~ p_UseFormIcon               - use forms icon                        (bool)
'/~ p_UseFormCaption            - use forms caption text                (bool)
'/~ p_TopSizingBorder           - top sizing border height              (long)
'/~ p_BottomSizingBorder        - bottom sizing border                  (long)
'/~ p_MinFormHeight             - minimum form height dimension         (long)
'/~ p_MinFormWidth              - minimum form width dimension          (long)
'/~ p_CaptionOffsetX            - caption offset position X             (long)
'/~ p_CaptionOffsetY            - caption offset position Y             (long)
'/~ p_ThreadAccel               - thread accleration switch             (bool)

'/~ Menu Properties /~
'/~ p_IMenuBg                   - menu bg image                         (img)
'/~ p_IMenuRollover             - menu hilite bar image                 (img)
'/~ p_IMenuBarBg                - menu caption bar image                (img)
'/~ p_MenuCustom                - use skinned menus                     (bool)
'/~ p_MenuActiveForeColor       - active menu item selected color       (long)
'/~ p_MenuInActiveForeColor     - active menu item color                (long)
'/~ p_MenuBackgroundColor       - menu bg color                         (long)
'/~ p_MenuOffsetX               - menu start position X                 (long)
'/~ p_MenuOffsetY               - menu start position Y                 (long)
'/~ p_MenuImageList             - menu image list                       (var)
'/~ p_MenuIconIndex             - icon index collection                 (long)
'/~ p_MenuRollOver              - custom caption rollover effect        (bool)
'/~ p_MenuRollOverColor         - captionbar rollover accent color      (long)
'/~ p_MenuFont                  - specify menu font (default is form)   (str)
'/~ p_MenuRollOverStyle         - rollover bar style                    (bool)
'/~ p_OfficeXpStyle             - office xp style menu                  (bool)
'/~ p_MenuTransparent           - use transparent menus (2K/XP)         (bool)
'/~ p_MenuTransIdx              - transparency level (1 - 255)          (byte)

'/~ Command Button Properties /~
'/~ p_ICmdImg                   - command button image                  (img)
'/~ p_SkinCommand               - skin command buttons                  (bool)
'/~ p_CmdTransColor             - transparent color                     (long)
'/~ p_CmdCaption                - control caption                       (str)
'/~ p_CmdRenderOffsetX          - stretchblt offset X                   (long)
'/~ p_CmdRenderOffsetY          - stretchblt offset Y                   (long)
'/~ p_CmdForeColor              - button text color                     (long)
'/~ p_CmdColorHiLite            - button over text color                (long)
'/~ p_CmdDrawFocused            - draw focus rect                       (bool)
'/~ p_CmdTextAntiAliased        - use anti-aliased caption              (bool)
'/~ p_CmdIconPosition           - icon position                         (enum)

'/~ Option Button Properties /~
'/~ p_IOptImg                   - option button image                   (img)
'/~ p_SkinOptionButton          - skin option buttons                   (bool)
'/~ p_OptTransColor             - transparent color                     (long)

'/~ CheckBox Properties /~
'/~ p_IChkImg                   - checkbox image                        (img)
'/~ p_SkinCheckbox              - skin checkboxes                       (bool)
'/~ p_ChkTransColor             - transparent color                     (long)

'/~ ComboBox Properties /~
'/~ p_IComboImg                 - combobox image                        (img)
'/~ p_SkinComboBox              - skin checkboxes                       (bool)
'/~ p_ComboTransColor           - transparent color                     (long)
'/~ p_ComboFrameColor           - base frame color                      (long)
'/~ p_ComboFrameHighLite        - frame accent color                    (long)
'/~ p_ComboFrameStyle           - frame style                           (enum)

'/~ WaterMark Properties /~
'/~ p_IWaterMark                - watermark image                       (img)
'/~ p_WaterMarkPosition         - position options                      (enum)
'/~ p_WaterMarkCtrlHnd          - control handle                        (long)

'/~ Frame Style Properties /~
'/~ p_FrameColorNormal          - frame normal color                    (long)
'/~ p_FrameColorHover           - frame mouse over color                (long)
'/~ p_FrameColorFocused         - frame focused color                   (long)
'/~ p_FrameListView             - apply to listview                     (bool)
'/~ p_FrameTextBox              - apply to textbox                      (bool)
'/~ p_FrameTreeView             - apply to treeview                     (bool)
'/~ p_FrameFileListBox          - apply to file listbox                 (bool)
'/~ p_FrameDirListBox           - apply to dir listbox                  (bool)
'/~ p_FrameListBox              - apply to listbox                      (bool)
'/~ p_FramePictureBox           - apply to picturebox                   (bool)

'/~ ListView Properties /~
'/~ p_SkinListView              - skin listview headers                 (bool)
'/~ p_OListViewObj              - control object                        (obj)
'/~ p_ListViewFlatPanel         - use flat header bg                    (bool)
'/~ p_ListViewTextAntiAliased   - use anti-aliased caption              (bool)
'/~ p_ListViewForeColor         - header font color                     (long)
'/~ p_ListViewHighLite          - header font highlite color            (long)

'/~ ScrollBar Properties /~
'/~ p_IListViewHeader           - listview header image                 (img)
'/~ p_IScrollVTrack             - scrollbar vertical track              (img)
'/~ p_IScrollVThumb             - scrollbar vertical thumb              (img)
'/~ p_IScrollVBtDwn             - scrollbar vertical button down        (img)
'/~ p_IScrollVBtUp              - scrollbar vertical button up          (img)
'/~ p_IScrollHTrack             - scrollbar horizontal track            (img)
'/~ p_IScrollHThumb             - scrollbar horizontal thumb            (img)
'/~ p_IScrollHBtLft             - scrollbar horizontal button left      (img)
'/~ p_IScrollHBtRgt             - scrollbar horizontal button right     (img)
'/~ p_IScrollSizer              - scrollbar sizer corner                (img)
'/~ p_ScrollBarAdd              - add scrollbars to a control           (long)

'/~ ProgressBar Properties /~
'/~ p_OControlObj               - parent object                         (obj)
'/~ p_IPBar                     - bar image                             (img)
'/~ p_ForeColor                 - bar forecolor (n/a)                   (long)
'/~ p_BackColor                 - inner bar backcolor                   (long)
'/~ p_BarColor                  - bar color                             (long)
'/~ p_BarGradColor              - gradient secondary color              (long)
'/~ p_BarStyle                  - bar style                             (enum)

'/~ Slider Properties /~
'/~ p_OParentObj                - parent object                         (obj)
'/~ p_ISldVTrack                - vertical track image                  (img)
'/~ p_ISldVThumb                - vertical thumb image                  (img)
'/~ p_ISldHTrack                - horizontal track image                (img)
'/~ p_ISldHThumb                - horizontal thumb image                (img)
'/~ p_UseToolTips               - use tooltips (n/a)                    (bool)
'/~ p_BackColor                 - bar backcolor                         (long)


'/-! Disclaimer
'/~ Obviously no warranty or liability, or any responsibility in any way imaginable, is expressed or implied.
'/~ Use this software in any way you wish under a relaxed GNU, but all responsibilities
'/~ are yours entirely. This example program is only a demonstration for educational purposes, if you
'/~ use these skins, you will need to obtain the original artists permission, (that is -your- responsibility).

'/-* Credits/Cudos
'/~ First off, hats off to Steve, without whom, this would not have been possible..
'/~ Uses a variation of Paul Catons excellent WinSubHook3 subclasser.
'/~ Used code ideas from many different places, I'm sure credits are due to Carles and Mario.

'/-> History
'/~ Well I finally ended up needing Neo for an app, so it was time for a serious rewrite..
'/~ Discarded much of the drawing routines from previous rev, and went from tiling w/ bitblt, to
'/~ using stretchblt and larger images, (inline with commercial skinning apps..). The result is an
'/~ engine that can, (with very little adjustment), handle skin images from common commercial
'/~ skinning apps. This version has included menus -with a twist-, I liked the way some of the skinned
'/~ menus looked in WB, so I thought I would implement it, I think it turned out good..

'/-> Things to Know for 2.7.3..
'/~ Command buttons must have their style property set to graphical.
'/~ Command buttons must also be situated [only] on a frame control.
'/~ An alternate method for skinning command buttons is demonstrated in the commandbutton
'/~ standalone project.

'/-! Compiling and Distribution !
'/~ When compiling a library for release, -always- set the version compatability to binary.
'/~ Project > Properties > Component > Version Compatabilty > 'Binary Compatability'.
'/~ If you plan to distribute compiled dll, Rename it! Use a unique name, with a app/company name prefix,
'/~ ex. sgrSkinLib.dll. If you do not, you are taking a risk that someone elses version will overwrite
'/~ your own, and your app will stop working!
'/~ The compile switches are set, but you could change them to your liking and test responsiveness
'/~ under different option settings.
'/~ If you use this in a commercial product, I expect to be notified, and a mention in the about/help or
'/~ other suitable spot is a basic proffesional courtesy, ex. 'Skinning engine courtesy of John Underhill of
'/~ NSPowertools.com.' Otherwise you are free to use this any way you like [excepting malicious].
'/~ You can distribute this project to other vb related websites, but headers must remain intact,
'/~ and proper credit must be given.
'/-> What I would like to see, [for .Net users], is some smart fellow, convert this project to vb.Net,
'/~ and publish it on the other side of psc. I have done quite a bit of searching, and could not find a single
'/~ example of a skinning library in vb.Net, (so you would be making a lot of people very happy!)
'/~ Don't wait for me to do it, I don't write in vb.Net, and I probably won't get around to porting this
'/~ to C# for quite a while.. (and when I do, it will probably go to sourceforge).

'/-> 98/ME users -
'/~ I don't claim this is legacy OS compliant, but have made an attempt at backwards availability.
'/~ If you are using a legacy OS, set the p_CustomCaption property to False. You will need to use
'/~ caption headers that are the same shape as a standard caption. The custom caption property
'/~ turns off transparency on frame skinning, transparentblt api leaks in 98/Me, so an alternate
'/~ routine 'TransMask' is used, but, blit masking is computationally expensive by comparison,
'/~ so setting this property to false will yield better render times.
'/~ Some controls require transparent regions, this is unavoidable. Keep their usage to a minimum.

'/-> Creating your own Skins
'/~ Before you begin, open frmMain, and resize the picMain picture so that the form background is exposed.
'/~ You will now see a number of images used in the 'lime' theme skin. Making a skin only takes some modest
'/~ drawing skills, and a bit of patience, with a little time, you should be able to sample and build a skin
'/~ in an hour or less..
'/-> Form: 9 images (CaptionBar/MenuBar/left side/right side/bottom/4 button groups)
'/~ You will note that areas that are to be made invisible use the magenta (&HFF00FF) color, but this is not
'/~ required. The top left pixel of an image is sampled and becomes the transparency color, unless specifically
'/~ overrided.
'/~ The forms frame consists of 3 sides and the caption bar, if you look at the images closely, you will see that
'/~ the finished corner edges are on the left and right side images. There are also 4 button groups (min/max/restore/close)
'/~ The button image states are (left to right) up/over/down. The frame and image buttons are mandatory objects, the
'/~ library instance will not load if any of these images are absent. Also note, all image properties begin with the
'/~ prefix 'p_I' so that they can be easily identified in the property dialog.
'/~ If you want to use menus on your form, you must include the p_IMenuBarBg image, or the menu will not reveal.
'/~ If you want to use an image background the p_IMenuBarBg and p_IMenuRollover must also be included.
'/-> CheckBox: 1 image
'/~ Image has 12 states, side by side, from left: 0-[empty] normal, 1-focused, 2-pressed, 3-disabled, 4-[value]
'/~ normal, 5-focused, 6-pressed, 7-disabled 8-[greyed] normal, 9-focused, 10-pressed, 11-disabled
'/-> ComboBox: 1 image
'/~ image has 4 states: 0-normal, 1-focus, 2-down, 3-disabled
'/-> Command Button: 1 image
'/~ Image has 5 states: 0-normal, 1-pressed, 2-disabled, 3-over, 4-focused
'/! Command buttons must have their style property set to 'graphical'
'/-> Listview Header: 1 image
'/~ image has 4 states: 0-normal, 1-focus, 2-down, 3-disabled
'/-> Option Button: 1 image
'/~ image has 8 states: 0-[empty] normal, 1-focused, 2-pressed, 3-disabled,
'/~ 4-[value] normal, 5-focused, 6-pressed, 7-disabled
'/-> Progress Bar: 1 image
'/~ 1 state, tick
'/-> Scroll Bars: 9 images (vert: arrowtop/arrowbottom/thumb/track horz: arrowleft/arrowright/thumb/track, sizer)
'/~ 2 state arrows
'/~ 2 sets of images and a sizer image. Both sets are required to load class.
'/-> Slider: 4 Images (vert: thumb/track horz: thumb/track)
'/~ 2 state thumb
'/~ I would suggest you simply look closely at the image groups on the form if in doubt. The image groupings
'/~ are also very similar to what is used in standard commercial skinning apps.

'/-# Version 2.1 - August 01, 2006
'/~ Fixed a number of small bugs related to menuing. Crash after skin switch/menu init resolved.

'/-# Version 2.2 - August 02, 2006
'/~ Finally fixed the menus keeping state after a skin change, all seems well now..

'/-# Version 2.3.2 - August 03, 2006
'/~ Engine is now 100% flicker free! (Abuse it, I dare ya! ;o)
'/~ Rendering now to a temp dc, then blitting whole skin in one go, provides for 0% flicker, (at last!)
'/~ Many other significant changes to this rev. Skinning format is 100% inline with WB4. That means,
'/~ you can now simply open up a skin editor, export frame/caption bar/buttons, and menu background
'/~ from your favorite skin, add them to pictureboxes or res file, and voila! Your app now looks
'/~ like that skin!
'/~ Engine also now supports irregular shaped forms, as is demonstrated in this example, so go nuts,
'/~ create apps with whatever style you like.. sky is the limit
'/~ If you use a skin, I strongly advise you to contact the author and get permission if you plan to publish
'/~ the app, it is a basic courtesy, and could avoid problems. Most of them will allow it, usually for
'/~ just a mention in the about of your product, (I know, I have asked, and that is typical). I also
'/~ expect a mention in the appropriate spot with a published product, ie 'RCM GII skinning engine courtesy
'/~ of John Underhill at NSPowertools.com'
'/~ Added a number of properties/methods, most notably, caption title frame image, menu bar image, and animated
'/~ image that resides on the captionbar.
'/~ Windows98/ME users will want to use p_CustomCaption set to false, as transparentblt is known to cause leaks,
'/~ what this means that cyan colored transparent areas will not be available on that older platform, so you
'/~ would need to stick to regular shaped forms like the lg-esque demo.
'/~ Reported crash on form close seems resolved, (asm string in subclasser), new version has the fix.
'/~ Flicker in client area also solved, (one too many window_redraw calls)

'/-# Version 2.4 - August 06, 2006
'/~ Fixed some more bugs. But most importantly, command and option buttons, are now skinned!
'/~ This is first volley at skinning it all, (one lib to draw them all, one lib to bind them..).
'/~ I think 2.5 will skin scrollbars, comboboxes, and listview headers. Would be nice if one library could
'/~ take the place of a dozen uc's, yes? Stay tuned, the best is yet to come..

'/-# Version 2.5 August 12, 2006
'/~ Added rendering support and check for 98/ME, should work ok now in old OS, (someone let me know).
'/~ Command buttons, option buttons, checkboxes, comboboxes, imagecombo, and drivelist are now also skinned.
'/~ Many small fixes and improvements to speed.
'/~ Auto samples transparency color (x 0, y 0), alpha image support with auto conversion from 24b to 32b.
'/~ Remember to compile this though, (is 70%+ faster compiled!). Might (probably) switch all image processing
'/~ api to a typelib for more pickup, I wish vb was not so slow ;o(.
'/~ Also added transparency, xp style, and many extended style options to menus.
'/~ Thats eight controls skinned, five to go.. (progress bar/statusbar/scrollbars/listview/slider)
'/-# Update: v. 2.5.2 August 13, 2006
'/~ Spotted a bug with the chkbox state on form refresh, so redid logic in checkbox and command buttons.
'/~ Added icon support to command buttons and re-uploaded. Had RemoveKeyboardHook remmed, (whoops!), fixed now..
'/-# Update: v. 2.5.3 August 14, 2006
'/~ Adjusted logic in option buttons, (hopefully for the last time ;o). Button reveal on caption down is fixed.
'/~ Form moveable when maximized is fixed. Blue frame visible when maximized/restored is fixed. Added a sub
'/~ 'Change_Caption' to change caption text on the fly.

'/-@ Bugfixes 2.5
'/~ Some of these were kind of interesting, so I thought I'd share..
'/~ * Option buttons - I thought this would be one of the simplest controls to modify, boy was I wrong!
'/~ With an option button array, the parent controls some of the message handling for the control members.
'/~ When you click down on a button, killfocus and paint messages are also sent to the last focused member.
'/~ What was happening was, when clicking option button to change the skin, messages were coming in while controls
'/~ subclass instance was being detatched, creating a circular reference. I tried using object pointers to resolve it,
'/~ but that made it worse! I finally just restructured logic in drawing routines, and this seems to have fixed it.
'/~ * Multithreading - this is very unstable in vb, and switch between skins was causing a gpf. I removed it for
'/~ now until a better solution could be found.
'/~ * Menu - Timer was not being destroyed, causing buildup of unresolved timers, (a clear indication of this is
'/~ when using 'With controlx' when you type the period, properties menu immediately shows then dissapears). Resolved
'/~ by allowing only one simultaneous timer during mouseover routine.
'/~ * Transparency - First off, transparentblt causes a massive gdi leak in 98/ME, so I moved all drawing routines
'/~ to a class wrapper, (clsRender), and use an OS check to choose between a blt mask and transparentblt api.
'/~ Class wrapper will use a typelib for all rendering in 2.5. Transparency on form was also skewed while form was
'/~ being resized, this is because stretchblt is not rendering nearest neighbor, but shifting colors of image,
'/~ causing transparent color to change. Solved this by optionaly sampling (x 0, y 0) pixel of image for transparent
'/~ color before rendering.
'/~ * Alt Key - Alt messages are now being consumed with a keyboard hook, menu acclerators should be added in 2.7.

'/-# Version 2.6.1 August 22, 2006
'/~ A lot of changes and additions to this version.. Three subclassing methods are demonstrated,
'/~ 1) subclassing an objects paint messages 2) using WM_DRAWITEM messages and struct as in command buttons
'/~ 3) using the NM_CUSTOMDRAW structure in listview headers
'/~ Each method has strengths and weaknesses, drawitem method will only work with certain controls, (if
'/~ you wanted to skin a toolbar for example, customdraw might be the best choice, while drawitem requires
'/~ control be set to graphical style at design time to work).
'/~ Skinned the ever-elusive listview header, (I could not find an example of this in vb anywhere), and
'/~ scrollbars are skinned across the control set, using a window on window method demonstrated by Carles.
'/~ Library can now do 70% of what wb4 can do, and if you wanted to, you could use methods demonstrated here
'/~ to take it the rest of the way, (but I would do it in C++, -vb is just too slow!).
'/~ Command buttons must have style set to graphical at design time, but drawitem method works well for them,
'/~ and they are now flicker free.
'/~ The bulk of the common control set is now subclassed, most of them use modified frame style and rollover effect,
'/~ but if you want to take that further, the groundwork has been laid, just add the paint/message routines to suit
'/~ your style goals.
'/~ Next rev will use a typelib for all the rendering api, should give it a little pickup. Will also likely
'/~ switch it all over to object pointers.

'/-# Version 2.6.2 August 23, 2006
'/~ Fixed listview headers, they are solid as a rock now. Changed scrollbars to blit in end pieces on thumbs,
'/~ now images are not distorted by stretching. Access to library has been changed. Controls now need to be
'/~ initiated through their respective classes. This change allows for a significant increase in performance
'/~ and has eliminated crashes when skin was swapped out, (due to pointer ref issues when redirecting through
'/~ a hosting class). The result is a dll that is much faster and more stable.
'/~ Reinstated thread accelerators to further improve performance.

'/-# Version 2.7 August 30, 2006
'/~ Complete rewite of library. Removed every unused constant and declaration, and trimmed message handling
'/~ to absolute minumum, with many optimizations throughout. Migrated api in subclasser and render class to
'/~ 'GPX Accelerator' typelib. Result is a substantial gain in performance and stability.
'/~ Added multi-style Progress bar, and slider controls to library.
'/~ Many small changes and improvements, like ellipses on listview header, and most of the glitches have
'/~ been found and resolved.
'/~ Largest overall problem with the library seems to be dealt with, that is gpf due to circular reference.
'/~ Note* if you change this, make absolutely certain that every single resource you create is properly disposed of.
'/~ Setting an activex library to nothing with an unresolved resource, (even as small as a pen, or selector object),
'/~ will very likely cause a gpf access violation. This is beacuse the dll will not unload with a unresolved objects
'/~ or existing child references. I dealt with this partially by creating 'soft references' to objects using pointers,
'/~ and being very careful about avoiding resource leaks.
'/~ The example form is running 36 simultaneous subclass instances, with no flicker, and great response times.
'/~ If you want a better idea of what that means, take your favorite button control and cover a form with
'/~ 36 instances of it, and compare the response. In the real world, it is unlikely you should have this many controls on
'/~ the same form, (bad UI design).
'/~ What it doesn't do.. it doesn't skin toolbars, statusbar or tab control, not because it is not possible,
'/~ but because I simply didn't get around to doing it, (takes forever to write this stuff ;o).
'/~ Certainly though, methods demonstrated here give you the insight needed to skin most any control, so the rest is up to you.
'/~ I included both versions with typelib, and without. Typelib version is superior, but I know some folks have
'/~ a phobia about additional references. I will also include all of the updated working projects, this way, you can
'/~ include the classes in your projects if you need only to skin a particular type of control.
'/~ Barring a couple inevitable bug fixes, this is the last version. 98/ME not supported. You want more features, add them.

'/-# Version 2.7.2 September 07, 2006
'/~ Swapped subclasser with faster winsubhook3 variation throughout. The result is a much faster and more stable library.
'/~ Typelib version implements all subclasser and common drawing api via typelib, and added IRunnable interface for
'/~ async processing. Methods in many peripheral controls changed to BeginPaint/EndPaint. This allowed for a finer
'/~ control over paint events, and less message hooking. Control gained meant progressbar and listview headers could be
'/~ painted directly without flicker, so child window overlays were removed. Most of the classes were then rewritten
'/~ with a mind to more compact and efficient code. The result is far faster and more stable then previous releases,
'/~ even with 34 subclassing instances, cpu footprint is small, and rendering is very fast.
'/~ Also added ole color support throughout, fixed slider bars, menus, progress bars, etc.
'/-# Version 2.7.2 and 1/2 -September 08, 2006
'/~ Added disabled shadow back to optionbutton and checkbox and fixed mouseover image. Fixed button reveal on caption
'/~ bar if form maximized then right most corner clicked. Made form unmoveable when maximized.
'/-# Version 2.7.2 and 3/4 -September 08, 2006
'/~ Form no longer needs to be reloaded with skin change, with one trade off.. command buttons must be on a frame control.

'/-@ Bugfixes 2.7
'/~ Gdi leaks, lots of them.. Hunted them down one by one, they are all resolved now. Crashes on skin change. This
'/~ was happening for a couple of reasons.. First, activex dll implementation -not so hot ;o) The gdi leaks were
'/~ causing library not to unload completely, so memory was still occupied by previous library instance which in turn
'/~ caused a gpf with access violation. Also at fault, previous subclasser did not always unload thunks before
'/~ reinstance, again causing gpf. Swapped subclasser and plugged the holes. Tested by switching skins 50 times in a row,
'/~ and opening closing app 50 times in a row, -no crash.
'/~ One small gdi leak somewhere I couldn't find, probably in watermark or storedc somewhere, but it is only 2 objects
'/~ per reload, so no big deal, (if you spot it though, let me know..).
'/~ Biggest leak was in progressbar, (2 objects per tick!), and in menu, (3 objects per show).

'/-@ Bugfixes 2.7.3
'/~ Option button redraw_all restored. vImageList test removed in menu loader. Disabled text re-added to checkbox
'/~ and option button. Fixed incorrect scroll thumb image in lg example.

'/-& Some Stray Thoughts
'/~ I remember seeing a comment on another post, to the effect that skinning with bitmaps was an inferior method to
'/~ custom drawing the control, (that is wrong on sooo many levels..). How do you think XP styles are being
'/~ rendered? For the most part, it is using bitmaps stored in a dll and blitting over controls with the image,
'/~ (ever notice distortion on inner border of xp style button if it is excessively long -stretching- a bitmap)
'/~ Aside from the fact that when you use a usercontrol, you are stuck with the authors styles, more importantly,
'/~ when you draw a control pixel by pixel, you are performing hundreds, (even thousands) of calculations to
'/~ create the effect -with every paint-, by comparison, a bitblock transfer is only one high level call..
'/~ Vb is not the best language to create a skinning engine, mostly because it is very slow at rendering graphics
'/~ due to api (mis)management, type conversions and threading issues. If you were going to write a winblinds style app,
'/~ C or C++ would be your best bet. But given that, this project still demonstrates most of the methods required to
'/~ skin almost any control, and now you have a way to create any style you want for your applications, all done
'/~ through a single interface, and without the need for resource expensive usercontrols.
'/~ The problem with usercontrols is that they are inherently ineffecient. They are trying to reproduce qualities
'/~ of standard controls that are written in C, which are much faster, more stable, and have withstood the test of time.
'/~ Usercontrols became popular because pre-vb5, there was no way to subclass a common control to manipulate
'/~ its style, but with the inclusion of the addressof operator, things changed for the better.
'/~ What I have done here just scratches the surface, and only touches upon methods using blitting and overpaints,
'/~ but there is so much more that is possible.. almost any style, property, action can be added to a control, just
'/~ with a bit of subclassing, some research into the methods, and a little imagination.. ~sky is the limit~

'-@ steppenwolfe_2000@yahoo.com
'-> Cheers,
'-> John


'/*
'/**
'/***
'/***   -Project Mission-
'/***   2.5.3    ***
'/***   fix checkbox/optbtn logic           -> done
'/***   command button icon                 -> done
'/***   2.6
'/***   skin listbox                        -> done
'/***   skin picturebox                     -> done
'/***   skin treeview                       -> done
'/***   skin textbox (w/ watermark)         -> done
'/***   frame styles                        -> done
'/***   skin scrollbars                     -> done
'/***   skin listview (w/ headers)          -> done
'/***   2.6.2  ***
'/***   stabilize listview headers          -> done
'/***   scale scrollbars                    -> done
'/***   2.7    ***
'/***   typelib
'/***   skin progressbar                    -> done
'/***   skin slider                         -> done
'/***   inactive form                       -> skipped
'/***   alt accelerators                    -> skipped
'/***
'/**
'/*


Implements WinSubHookGX.GXISubclass

Private Const GWL_STYLE                 As Long = (-16)
Private Const GWL_EXSTYLE               As Long = (-20)
Private Const DI_NORMAL                 As Long = &H1 Or &H2
Private Const FW_NORMAL                 As Long = 400
Private Const FW_BOLD                   As Long = 700
Private Const LOGPIXELSY                As Long = 90
Private Const DST_COMPLEX               As Long = &H0
Private Const DST_TEXT                  As Long = &H1
Private Const DST_PREFIXTEXT            As Long = &H2
Private Const DST_ICON                  As Long = &H3
Private Const DST_BITMAP                As Long = &H4
Private Const DSS_NORMAL                As Long = &H0
Private Const DSS_UNION                 As Long = &H10
Private Const DSS_DISABLED              As Long = &H20
Private Const DSS_MONO                  As Long = &H80
Private Const DSS_RIGHT                 As Long = &H8000
Private Const DCX_WINDOW                As Long = &H0&
Private Const DCX_WINDOWX               As Long = &H1&
Private Const DCX_CACHE                 As Long = &H2&
Private Const DCX_NORESETATTRS          As Long = &H4&
Private Const DCX_CLIPCHILDREN          As Long = &H8&
Private Const DCX_CLIPSIBLINGS          As Long = &H10&
Private Const DCX_PARENTCLIP            As Long = &H20&
Private Const DCX_EXCLUDERGN            As Long = &H40&
Private Const DCX_INTERSECTRGN          As Long = &H80&
Private Const DCX_EXCLUDEUPDATE         As Long = &H100&
Private Const DCX_INTERSECTUPDATE       As Long = &H200&
Private Const DCX_LOCKWINDOWUPDATE      As Long = &H400&
Private Const DCX_VALIDATE              As Long = &H200000
'/* font quality
Private Const DEFAULT_QUALITY           As Long = &H0
Private Const DRAFT_QUALITY             As Long = &H1
Private Const PROOF_QUALITY             As Long = &H2
Private Const NONANTIALIASED_QUALITY    As Long = &H3
Private Const ANTIALIASED_QUALITY       As Long = &H4
Private Const WS_EX_LAYERED             As Long = &H80000
Private Const WS_EX_TRANSPARENT         As Long = &H20&
Private Const LWA_ALPHA                 As Long = &H2&
'/* redraw
Private Const RDW_ALLCHILDREN           As Long = &H80
Private Const RDW_ERASE                 As Long = &H4
Private Const RDW_ERASENOW              As Long = &H200
Private Const RDW_FRAME                 As Long = &H400
Private Const RDW_INTERNALPAINT         As Long = &H2
Private Const RDW_INVALIDATE            As Long = &H1
Private Const RDW_NOCHILDREN            As Long = &H40
Private Const RDW_NOERASE               As Long = &H20
Private Const RDW_NOFRAME               As Long = &H800
Private Const RDW_NOINTERNALPAINT       As Long = &H10
Private Const RDW_UPDATENOW             As Long = &H100
Private Const RDW_VALIDATE              As Long = &H8
Private Const RDW_NRPGROUP              As Long = RDW_INVALIDATE Or _
    RDW_ERASE Or RDW_UPDATENOW Or RDW_FRAME

'/* system metrics
Private Enum SYSTEM_METRICS
    SM_CXSCREEN = 0
    SM_CYSCREEN = 1
    SM_CXVSCROLL = 2
    SM_CYHSCROLL = 3
    SM_CYCAPTION = 4
    SM_CXBORDER = 5
    SM_CYBORDER = 6
    SM_CYVTHUMB = 9
    SM_CXHTHUMB = 10
    SM_CXICON = 11
    SM_CYICON = 12
    SM_CXCURSOR = 13
    SM_CYCURSOR = 14
    SM_CYMENU = 15
    SM_CXFULLSCREEN = 16
    SM_CYFULLSCREEN = 17
    SM_CYKANJIWINDOW = 18
    SM_MOUSEPRESENT = 19
    SM_CYVSCROLL = 20
    SM_CXHSCROLL = 21
    SM_CXMIN = 28
    SM_CYMIN = 29
    SM_CXSIZE = 30
    SM_CYSIZE = 31
    SM_CXFRAME = 32
    SM_CYFRAME = 33
    SM_CXMINTRACK = 34
    SM_CYMINTRACK = 35
    SM_CYSMCAPTION = 51
    SM_CXMINIMIZED = 57
    SM_CYMINIMIZED = 58
    SM_CXMAXTRACK = 59
    SM_CYMAXTRACK = 60
    SM_CXMAXIMIZED = 61
    SM_CYMAXIMIZED = 62
End Enum

Private Enum SWP_FLAGS
    SWP_NOSIZE = &H1
    SWP_NOMOVE = &H2
    SWP_NOZORDER = &H4
    SWP_NOREDRAW = &H8
    SWP_NOACTIVATE = &H10
    SWP_FRAMECHANGED = &H20
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_NOCOPYBITS = &H100
    SWP_NOOWNERZORDER = &H200
    SWP_NOSENDCHANGING = &H400
End Enum

Public Enum ECHIGHLIGHTCONSTEX
    ECSTANDARDEX = 0
    ECGRADIENTEX = 1
    ECBUTTONEX = 2
    ECROLLOVEREX = 3
End Enum

Private Enum WD_STATE
    WD_MINIMIZED = 1
    WD_NORMAL = 2
    WD_MAXIMIZED = 3
End Enum

Private Enum ECNBUTTONSTATES
    ECN_UP = &H0
    ECN_OVER = &H1
    ECN_DOWN = &H2
End Enum

Public Enum ECNCMD
    SC_ARRANGE = &HF110&
    SC_CLOSE = &HF060&
    SC_MAXIMIZE = &HF030&
    SC_MINIMIZE = &HF020&
    SC_MOVE = &HF010&
    SC_NEXTWINDOW = &HF040&
    SC_PREVWINDOW = &HF050&
    SC_RESTORE = &HF120&
    SC_SIZE = &HF000&
End Enum

Public Enum EHITCONSTANTS
    HTBORDER = 18
    HTBOTTOM = 15
    HTBOTTOMLEFT = 16
    HTBOTTOMRIGHT = 17
    HTCAPTION = 2
    HTCLIENT = 1
    HTGROWBOX = 4
    HTHSCROLL = 6
    HTLEFT = 10
    HTMAXBUTTON = 9
    HTMENU = 5
    HTMINBUTTON = 8
    HTNOWHERE = 0
    HTRIGHT = 11
    HTSYSMENU = 3
    HTTOP = 12
    HTTOPLEFT = 13
    HTTOPRIGHT = 14
    HTVSCROLL = 7
End Enum

Public Enum ICON_POSITIONEX
    IconLeft = &H1
    IconRight = &H2
    IconTop = &H3
End Enum


Private Type VERSION_INFO
    dwOSVersionInfoSize                 As Long
    dwMajorVersion                      As Long
    dwMinorVersion                      As Long
    dwBuildNumber                       As Long
    dwPlatformId                        As Long
    szCSDVersion                        As String * 128
End Type

Private Type LOGBRUSH
    lbStyle                             As Long
    lbColor                             As Long
    lbHatch                             As Long
End Type

Private Type POINTAPI
    x                                   As Long
    y                                   As Long
End Type

Private Type RECT
    left                                As Long
    top                                 As Long
    Right                               As Long
    Bottom                              As Long
End Type

Private Type WINDOWPOS
    hwnd                                As Long
    hWndInsertAfter                     As Long
    x                                   As Long
    y                                   As Long
    cx                                  As Long
    cy                                  As Long
    flags                               As Long
End Type

Private Type NCCALCSIZE_PARAMS
    rgrc(0 To 2)                        As RECT
    lppos                               As Long
End Type

Private Type BITMAP
    bmType                              As Long
    bmWidth                             As Long
    bmHeight                            As Long
    bmWidthBytes                        As Long
    bmPlanes                            As Integer
    bmBitsPixel                         As Integer
    bmBits                              As Long
End Type

Private Type LOGFONT
    lfHeight                            As Long
    lfWidth                             As Long
    lfEscapement                        As Long
    lfOrientation                       As Long
    lfWeight                            As Long
    lfItalic                            As Byte
    lfUnderline                         As Byte
    lfStrikeOut                         As Byte
    lfCharSet                           As Byte
    lfOutPrecision                      As Byte
    lfClipPrecision                     As Byte
    lfQuality                           As Byte
    lfPitchAndFamily                    As Byte
    lfFaceName                          As String * 32
End Type

Private Type BLENDFUNCTION
    BlendOp                             As Byte
    BlendFlags                          As Byte
    SourceConstantAlpha                 As Byte
    AlphaFormat                         As Byte
End Type

Private Type MINMAXINFO
    ptReserved                          As POINTAPI
    ptMaxSize                           As POINTAPI
    ptMaxPosition                       As POINTAPI
    ptMinTrackSize                      As POINTAPI
    ptMaxTrackSize                      As POINTAPI
End Type


Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, _
                                                                                         ByVal lpsz As String, _
                                                                                         ByVal cbString As Long, _
                                                                                         lpSize As POINTAPI) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, _
                                                       ByVal hwndTo As Long, _
                                                       lppt As Any, _
                                                       ByVal cPoints As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, _
                                                    ByVal y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, _
                                               ByVal hMenu As Long) As Long

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSrc As Any, _
                                                                     ByVal ByteLen As Long)

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        ByVal lParam As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Any) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long

Private Declare Function PtInRect Lib "user32" (lpRect As RECT, _
                                                ByVal x As Long, _
                                                ByVal y As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, _
                                                    lprcUpdate As Any, _
                                                    ByVal hrgnUpdate As Long, _
                                                    ByVal fuRedraw As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long) As Long

Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                   ByVal nCount As Long, _
                                                                   lpObject As Any) As Long

Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, _
                                                    ByVal nBkMode As Long) As Long

Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, _
                                                                      ByVal lpStr As String, _
                                                                      ByVal nCount As Long, _
                                                                      ByRef lpRect As RECT, _
                                                                      ByVal wFormat As Long) As Long

Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, _
                                                      ByVal xLeft As Long, _
                                                      ByVal yTop As Long, _
                                                      ByVal hIcon As Long, _
                                                      ByVal cxWidth As Long, _
                                                      ByVal cyWidth As Long, _
                                                      ByVal istepIfAniCur As Long, _
                                                      ByVal hbrFlickerFreeDraw As Long, _
                                                      ByVal diFlags As Long) As Long

Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, _
                                                       ByVal crColor As Long) As Long

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal nIndex As Long) As Long

Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, _
                                                ByVal nNumerator As Long, _
                                                ByVal nDenominator As Long) As Long

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal hRgn As Long) As Long

Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, _
                                                      ByVal x1 As Long, _
                                                      ByVal y1 As Long, _
                                                      ByVal x2 As Long, _
                                                      ByVal y2 As Long) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
                                                               ByVal HPALETTE As Long, _
                                                               pccolorref As Long) As Long

Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long, _
                                                  ByVal nWidth As Long, _
                                                  ByVal nHeight As Long, _
                                                  ByVal bRepaint As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersion As VERSION_INFO) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, _
                                                                  ByVal crey As Byte, _
                                                                  ByVal bAlpha As Byte, _
                                                                  ByVal dwFlags As Long) As Long

Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, _
                                             ByVal x As Long, _
                                             ByVal y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal dwRop As Long) As Long

Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private m_bCanSize                      As Boolean
Private m_bCanClose                     As Boolean
Private m_bCustomButtonPosition         As Boolean
Private m_bMinimise                     As Boolean
Private m_bMouseDownMinimise            As Boolean
Private m_bMouseDownMaximise            As Boolean
Private m_bMouseDownClose               As Boolean
Private m_bBtnChnge                     As Boolean
Private m_bUseFormCaption               As Boolean
Private m_bUseFormIcon                  As Boolean
Private m_bGThreadAccl                  As Boolean
Private m_bMaximized                    As Boolean
Private m_bWin32                        As Boolean
Private m_bFadeOutEffect                As Boolean
Private m_bCustomCaption                As Boolean
Private m_lCustomStartPos               As Long
Private m_lCustomEndPos                 As Long

Private m_lMinFormHeight                As Long
Private m_lMinFormWidth                 As Long

Private m_lIconDc                       As Long
Private m_lIconBmp                      As Long
Private m_lIconBmpOld                   As Long

Private m_lHostDc                       As Long

Private m_lHostHwnd                     As Long
Private m_CaptionFntClr                 As Long
Private m_lCaptionOffsetY               As Long
Private m_lCaptionOffsetX               As Long
Private m_bCenterCaption                As Boolean
Private m_bCaptionFrame                 As Boolean
Private m_lBottomSizingBorder           As Long
Private m_lTopSizingBorder              As Long

Private m_lButtonWidth                  As Long
Private m_lButtonHeight                 As Long
Private m_lControlOffsetX               As Long
Private m_lControlOffsetY               As Long

Private m_lBottomMetric                 As Long
Private m_lLeftMetric                   As Long
Private m_lRightMetric                  As Long
Private m_lTopMetric                    As Long
Private m_lCaptionMetric                As Long
Private m_tSizeState                    As RECT

Private m_lBottomBorderHeight           As Long
Private m_lBottomBorderWidth            As Long
Private m_lLeftBorderHeight             As Long
Private m_lLeftBorderWidth              As Long
Private m_lRightBorderHeight            As Long
Private m_lRightBorderWidth             As Long
Private m_lTopBorderHeight              As Long
Private m_lTopBorderWidth               As Long
Private m_lMenuBarBgHeight              As Long
Private m_lMenuBarBgWidth               As Long

Private m_lLastPosTop                   As Long
Private m_lLastPosLeft                  As Long
Private m_lCapIcon                      As Long
Private m_lDrawDc                       As Long

Private m_lMenuActiveForeColor          As Long
Private m_lMenuInActiveForeColor        As Long
Private m_lMenuBackgroundColor          As Long
Private m_lMenuOffsetX                  As Long
Private m_lMenuOffsetY                  As Long
Private m_sCapText                      As String
Private m_fntMenu                       As IFont
Private m_tMenuR                        As RECT
Private m_tBtn(0 To 2)                  As RECT

Private m_GSubclass                     As GXMSubclass
Private m_SizeInfo                      As MINMAXINFO
Private m_eLastHT                       As EHITCONSTANTS

Private m_cCaption                      As clsStoreDc
Private m_cLeftBorder                   As clsStoreDc
Private m_cRightBorder                  As clsStoreDc
Private m_cBottomBorder                 As clsStoreDc
Private m_cMinimize                     As clsStoreDc
Private m_cMaximize                     As clsStoreDc
Private m_cClose                        As clsStoreDc
Private m_cRestore                      As clsStoreDc
Private m_cMenuBarBg                    As clsStoreDc
Private m_cCaptionFrame                 As clsStoreDc

Private m_cMenu                         As clsRCMMenu
Private m_pCaption                      As StdPicture
Private m_pCCapFrame                    As StdPicture
Private m_pCLeft                        As StdPicture
Private m_pCRight                       As StdPicture
Private m_pCBottom                      As StdPicture
Private m_pCBoxMin                      As StdPicture
Private m_pCBoxMax                      As StdPicture
Private m_pCBoxRst                      As StdPicture
Private m_pCBoxCls                      As StdPicture
Private m_pCMenuBarBg                   As StdPicture
Private m_pCMenuBgPicture               As StdPicture
Private m_pCMenuRollover                As StdPicture
Private m_cRender                       As clsRender
'/* menu specific
Private m_bMenuActive                   As Boolean
Private m_bCustomMenu                   As Boolean
Private m_bMenuBar                      As Boolean
Private m_bMenuBg                       As Boolean
Private m_bOfficeXpStyle                As Boolean
Private m_bMenuTransparent              As Boolean
Private m_bMenuCustRollOver             As Boolean
Private m_lMenuRollOverColor            As Long
Private m_btMenuTransIdx                As Byte
Private m_colIcons                      As Collection
Private m_vImageList                    As Variant
Private m_StdMenuFont                   As StdFont
Private m_eMenuButtonStyle              As ECHIGHLIGHTCONSTEX
Private m_lParentObj                    As Long


'/* parent form
Public Property Get p_OParentObj() As Object

Dim oTemp   As Object

    If m_lParentObj = 0 Then
        Set p_OParentObj = Nothing
    Else
        CopyMemory oTemp, m_lParentObj, 4
        Set p_OParentObj = oTemp
        CopyMemory oTemp, 0&, 4
    End If

End Property

Public Property Set p_OParentObj(ByVal PropVal As Object)
    m_lParentObj = ObjPtr(PropVal)
End Property


'\\\\\\\\\\\\\\\\\\\\\\\\\-Menu Properties-/////////////////////////

Public Property Get p_IMenuBarBg() As StdPicture
'/* menu bar bg
    Set p_IMenuBarBg = m_pCMenuBarBg
End Property

Public Property Set p_IMenuBarBg(PropVal As StdPicture)
    Set m_pCMenuBarBg = PropVal
End Property

Public Property Get p_IMenuBg() As StdPicture
    Set p_IMenuBg = m_pCMenuBgPicture
End Property

Public Property Set p_IMenuBg(ByRef PropVal As StdPicture)

    If Not PropVal Is Nothing Then
        Set m_pCMenuBgPicture = PropVal
        m_bMenuBg = True
    Else
        m_bMenuBg = False
    End If

End Property

Public Property Get p_IMenuRollover() As StdPicture
    Set p_IMenuRollover = m_pCMenuRollover
End Property

Public Property Set p_IMenuRollover(ByRef PropVal As StdPicture)

    If Not PropVal Is Nothing Then
        Set m_pCMenuRollover = PropVal
        m_bMenuBar = True
    Else
        m_bMenuBar = False
    End If
    
End Property

Public Property Set p_MenuImageList(ByRef PropVal As Variant)

    If Not PropVal Is Nothing Then
        Set m_vImageList = PropVal
    End If

End Property

Public Property Let p_MenuIconIndex(ByVal sCaption As String, _
                                    ByVal nIndex As Long)

'/* store icon index data

On Error Resume Next

    If m_colIcons Is Nothing Then
        Set m_colIcons = New Collection
    End If
    
    If nIndex = -1 Then
        m_colIcons.Remove sCaption
    Else
        m_colIcons.Item sCaption
        If Err.Number = 0 Then
            m_colIcons.Remove sCaption
        End If
        m_colIcons.Add nIndex, sCaption
    End If

On Error GoTo 0

End Property

Public Property Get p_MenuActiveForeColor() As Long
'/* menu active item forecolor
    p_MenuActiveForeColor = m_lMenuActiveForeColor
End Property

Public Property Let p_MenuActiveForeColor(ByVal PropVal As Long)
    m_lMenuActiveForeColor = TranslateColor(PropVal)
    If Not m_cMenu Is Nothing Then
        m_cMenu.p_ActiveMenuForeColor = m_lMenuActiveForeColor
    End If
End Property

Public Property Get p_MenuInActiveForeColor() As Long
'/* menu item forecolor
    p_MenuInActiveForeColor = m_lMenuInActiveForeColor
End Property

Public Property Let p_MenuInActiveForeColor(ByVal PropVal As Long)
    m_lMenuInActiveForeColor = TranslateColor(PropVal)
    If Not m_cMenu Is Nothing Then
        m_cMenu.p_InActiveMenuForeColor = m_lMenuInActiveForeColor
    End If
End Property

Public Property Get p_MenuBackgroundColor() As Long
'/* menu background color
    p_MenuBackgroundColor = m_lMenuBackgroundColor
End Property

Public Property Let p_MenuBackgroundColor(ByVal PropVal As Long)
    m_lMenuBackgroundColor = TranslateColor(PropVal)
    If Not m_cMenu Is Nothing Then
        m_cMenu.p_MenuBackgroundColor = m_lMenuBackgroundColor
    End If
End Property

Public Property Get p_MenuCustom() As Boolean
'/* enable menu skinning
    p_MenuCustom = m_bCustomMenu
End Property

Public Property Let p_MenuCustom(ByVal PropVal As Boolean)
    m_bCustomMenu = PropVal
End Property

Public Property Get p_MenuRollOver() As Boolean
'/* custom caption rollover effect
    p_MenuRollOver = m_bMenuCustRollOver
End Property

Public Property Let p_MenuRollOver(ByVal PropVal As Boolean)
    m_bMenuCustRollOver = PropVal
    If Not m_cMenu Is Nothing Then
        m_cMenu.p_CustomRollOver = m_bMenuCustRollOver
    End If
End Property

Public Property Get p_MenuRollOverColor() As Long
'/* custom caption rollover effect
    p_MenuRollOverColor = m_lMenuRollOverColor
End Property

Public Property Let p_MenuRollOverColor(ByVal PropVal As Long)
    m_lMenuRollOverColor = TranslateColor(PropVal)
    If Not m_cMenu Is Nothing Then
        m_cMenu.p_RollOverColor = m_lMenuRollOverColor
    End If
End Property

Public Property Get p_MenuFont() As StdFont
'/* custom menu font
    Set p_MenuFont = m_StdMenuFont
End Property

Public Property Set p_MenuFont(ByVal PropVal As StdFont)
    Set m_StdMenuFont = PropVal
End Property

Public Property Get p_MenuOffsetX() As Long
'/* offset x
    p_MenuOffsetX = m_lMenuOffsetX
End Property

Public Property Let p_MenuOffsetX(ByVal PropVal As Long)
    m_lMenuOffsetX = PropVal
End Property

Public Property Get p_MenuOffsetY() As Long
'/* inactive color
    p_MenuOffsetY = m_lMenuOffsetY
End Property

Public Property Let p_MenuOffsetY(ByVal PropVal As Long)
    m_lMenuOffsetY = PropVal
    If Not m_cMenu Is Nothing Then
        m_cMenu.p_MenuOffsetY = m_lMenuOffsetY
    End If
End Property

Public Property Get p_OfficeXpStyle() As Boolean
'/* use office menu style
    p_OfficeXpStyle = m_bOfficeXpStyle
End Property

Public Property Let p_OfficeXpStyle(ByVal PropVal As Boolean)
    m_bOfficeXpStyle = PropVal
    If Not m_cMenu Is Nothing Then
        m_cMenu.p_OfficeXpStyle = m_bOfficeXpStyle
    End If
End Property

Public Property Get p_MenuTransparent() As Boolean
'/* use office menu style
    p_MenuTransparent = m_bMenuTransparent
End Property

Public Property Let p_MenuTransparent(ByVal PropVal As Boolean)
    m_bMenuTransparent = PropVal
    If Not m_cMenu Is Nothing Then
        m_cMenu.p_Transparent = m_bMenuTransparent
    End If
End Property

Public Property Get p_MenuTransIdx() As Byte
'/* inactive color
    p_MenuTransIdx = m_btMenuTransIdx
End Property

Public Property Let p_MenuTransIdx(ByVal PropVal As Byte)
    m_btMenuTransIdx = PropVal
    If Not m_cMenu Is Nothing Then
        m_cMenu.p_TransIndex = m_btMenuTransIdx
    End If
End Property

Public Property Get p_MenuRollOverStyle() As ECHIGHLIGHTCONSTEX
'/* inactive color
    p_MenuRollOverStyle = m_eMenuButtonStyle
End Property

Public Property Let p_MenuRollOverStyle(ByVal PropVal As ECHIGHLIGHTCONSTEX)
    m_eMenuButtonStyle = PropVal
End Property


'\\\\\\\\\\\\\\\\\\\\\\\\\-Form Properties-/////////////////////////

Public Property Get p_FadeOutEffect() As Boolean
'/* fade form on close
    p_FadeOutEffect = m_bFadeOutEffect
End Property

Public Property Let p_FadeOutEffect(PropVal As Boolean)
    m_bFadeOutEffect = PropVal
End Property

Public Property Get p_BottomSizingBorder() As Long
'/* bottom sizing border
    p_BottomSizingBorder = m_lBottomSizingBorder
End Property

Public Property Let p_BottomSizingBorder(ByVal PropVal As Long)
    m_lBottomSizingBorder = PropVal
End Property

Public Property Get p_ButtonOffsetX() As Long
'/* control button offset horz
    p_ButtonOffsetX = m_lControlOffsetX
End Property

Public Property Let p_ButtonOffsetX(ByVal PropVal As Long)
    m_lControlOffsetX = PropVal
End Property

Public Property Get p_ButtonOffsetY() As Long
'/* control button offset vert
    p_ButtonOffsetY = m_lControlOffsetY
End Property

Public Property Let p_ButtonOffsetY(ByVal PropVal As Long)
    m_lControlOffsetY = PropVal
End Property

Private Property Get p_CanClose() As Boolean
'/* test if window can close
Dim lStyle As Long
    lStyle = GetWindowLong(m_lHostHwnd, (-16))
    p_CanClose = ((lStyle And &H80000) = &H80000)
End Property

Private Property Get p_CanSize() As Boolean
'/* test if window is sizeable
Dim lStyle As Long
    lStyle = GetWindowLong(m_lHostHwnd, (-16))
    p_CanSize = ((lStyle And &H40000) = &H40000)
End Property

Public Property Get p_CaptionFntClr() As Long
'/* caption font color
    p_CaptionFntClr = m_CaptionFntClr
End Property

Public Property Let p_CaptionFntClr(ByVal PropVal As Long)
    m_CaptionFntClr = PropVal
End Property

Public Property Get p_CenterCaption() As Boolean
'/* caption shape offset start
    p_CenterCaption = m_bCenterCaption
End Property

Public Property Let p_CenterCaption(ByVal PropVal As Boolean)
    m_bCenterCaption = PropVal
End Property

Public Property Get p_CaptionFrame() As Boolean
'/* use caption text frame image
    p_CaptionFrame = m_bCaptionFrame
End Property

Public Property Let p_CaptionFrame(ByVal PropVal As Boolean)
    m_bCaptionFrame = PropVal
End Property

Public Property Get p_CaptionOffsetY() As Long
'/* caption text offset Y
    p_CaptionOffsetY = m_lCaptionOffsetY
End Property

Public Property Let p_CaptionOffsetY(ByVal PropVal As Long)
    m_lCaptionOffsetY = PropVal
End Property

Public Property Get p_CaptionOffsetX() As Long
'/* caption text offset X
    p_CaptionOffsetX = m_lCaptionOffsetX
End Property

Public Property Let p_CaptionOffsetX(ByVal PropVal As Long)
    m_lCaptionOffsetX = PropVal
End Property

Public Property Get p_ControlButtonPosition() As Boolean
'/* use custom button offsets
    p_ControlButtonPosition = m_bCustomButtonPosition
End Property

Public Property Let p_ControlButtonPosition(ByVal PropVal As Boolean)
    m_bCustomButtonPosition = PropVal
End Property

Public Property Get p_CustomStartPos() As Long
'/* caption shape offset start
    p_CustomStartPos = m_lCustomStartPos
End Property

Public Property Let p_CustomStartPos(ByVal PropVal As Long)
    m_lCustomStartPos = PropVal
End Property

Public Property Get p_CustomEndPos() As Long
'/* caption shape offset end
    p_CustomStartPos = m_lCustomEndPos
End Property

Public Property Let p_CustomEndPos(ByVal PropVal As Long)
    m_lCustomEndPos = PropVal
End Property

'~ form options ~
Public Property Get p_CustomCaption() As Boolean
'/* use transparency rendering
    p_CustomCaption = m_bCustomCaption
End Property

Public Property Let p_CustomCaption(ByVal PropVal As Boolean)
    m_bCustomCaption = PropVal
End Property

'~ form options ~
Public Property Get p_UseFormIcon() As Boolean
'/* use caption icon from form
    p_UseFormIcon = m_bUseFormIcon
End Property

Public Property Let p_UseFormIcon(ByVal PropVal As Boolean)
    m_bUseFormIcon = PropVal
End Property

Public Property Get p_UseFormCaption() As Boolean
'/* use caption text from form
    p_UseFormCaption = m_bUseFormCaption
End Property

Public Property Let p_UseFormCaption(ByVal PropVal As Boolean)
    m_bUseFormCaption = PropVal
End Property

Public Property Get p_ICaption() As StdPicture
'/* caption bar image
    Set p_ICaption = m_pCaption
End Property

Public Property Set p_ICaption(PropVal As StdPicture)
    Set m_pCaption = PropVal
End Property

Public Property Get p_ICBottom() As StdPicture
'/* control box close buttons
    Set p_ICBottom = m_pCBottom
End Property

Public Property Set p_ICBottom(PropVal As StdPicture)
    Set m_pCBottom = PropVal
End Property

Public Property Get p_ICBoxCls() As StdPicture
'/* control box close buttons
    Set p_ICBoxCls = m_pCBoxCls
End Property

Public Property Set p_ICBoxCls(PropVal As StdPicture)
    Set m_pCBoxCls = PropVal
End Property

Public Property Get p_ICBoxMax() As StdPicture
'/* control box maximize buttons
    Set p_ICBoxMax = m_pCBoxMax
End Property

Public Property Set p_ICBoxMax(PropVal As StdPicture)
    Set m_pCBoxMax = PropVal
End Property

Public Property Get p_ICBoxMin() As StdPicture
'/* control box minimize buttons
    Set p_ICBoxMin = m_pCBoxMin
End Property

Public Property Set p_ICBoxMin(PropVal As StdPicture)
    Set m_pCBoxMin = PropVal
End Property

Public Property Get p_ICBoxRst() As StdPicture
'/* control box restore buttons
    Set p_ICBoxRst = m_pCBoxRst
End Property

Public Property Set p_ICBoxRst(PropVal As StdPicture)
    Set m_pCBoxRst = PropVal
End Property

Public Property Get p_ICCapFrame() As StdPicture
'/* custom caption frame
    Set p_ICCapFrame = m_pCCapFrame
End Property

Public Property Set p_ICCapFrame(PropVal As StdPicture)
    Set m_pCCapFrame = PropVal
End Property

Public Property Get p_ICLeft() As StdPicture
'/* control box close buttons
    Set p_ICLeft = m_pCLeft
End Property

Public Property Set p_ICLeft(PropVal As StdPicture)
    Set m_pCLeft = PropVal
End Property

Public Property Get p_ICRight() As StdPicture
'/* control box close buttons
    Set p_ICRight = m_pCRight
End Property

Public Property Set p_ICRight(PropVal As StdPicture)
    Set m_pCRight = PropVal
End Property

Public Property Get p_MinFormHeight() As Long
'/* minimum form height
    p_MinFormHeight = m_lMinFormHeight
End Property

Public Property Let p_MinFormHeight(ByVal PropVal As Long)
    m_lMinFormHeight = PropVal
End Property

Public Property Get p_MinFormWidth() As Long
'/* minimum form width
    p_MinFormWidth = m_lMinFormWidth
End Property

Public Property Let p_MinFormWidth(ByVal PropVal As Long)
    m_lMinFormWidth = PropVal
End Property

'/* thread acceleration
Public Property Get p_ThreadAccel() As Boolean
    p_ThreadAccel = m_bGThreadAccl
End Property

Public Property Let p_ThreadAccel(ByVal PropVal As Boolean)
    m_bGThreadAccl = PropVal
End Property

Public Property Get p_TopSizingBorder() As Long
'/* top sizing border height
    p_TopSizingBorder = m_lTopSizingBorder
End Property

Public Property Let p_TopSizingBorder(ByVal PropVal As Long)
    m_lTopSizingBorder = PropVal
End Property

Private Sub Class_Initialize()
    InitCommonControls
    m_bWin32 = Compatability_Check
End Sub

Private Function Compatability_Check() As Boolean

Dim tVer  As VERSION_INFO

    tVer.dwOSVersionInfoSize = Len(tVer)
    GetVersionEx tVer
    If tVer.dwMajorVersion >= 5 Then
        Compatability_Check = True
    End If

End Function

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object

Dim oThis As Object

    CopyMemory oThis, lPtr, 4
    Set ObjectFromPtr = oThis
    CopyMemory oThis, 0&, 4

End Property

Private Property Get PtrFromObject(ByRef oThis As Object) As Long
    PtrFromObject = ObjPtr(oThis)
End Property


'/////////////////////////
'////   Entry Point  /////
'/////////////////////////

Public Sub Attach()
'/* initialization core

Dim tBmp    As BITMAP
Dim tObj    As Object
Dim lObj    As Long

'On Error Resume Next
    
    '/* test for base images
    If (p_ICaption Is Nothing) Or (p_ICBottom Is Nothing) Or _
        (p_ICLeft Is Nothing) Or (p_ICRight Is Nothing) Or _
        (p_ICBoxMin Is Nothing) Or (p_ICBoxMax Is Nothing) Or _
        (p_ICBoxRst Is Nothing) Or (p_ICBoxCls Is Nothing) Then
        Exit Sub
    End If
    
    '/* thread acceleration
    If m_bGThreadAccl Then
        Dim lThread As Long
        lThread = GetCurrentThreadId
        If Not lThread <= 0 Then
            mControl.ThreadAccelerate lThread, Thread_Maximum
        End If
    End If
    
    '/* if we are alive, reset
    If p_OParentObj.hwnd = m_lHostHwnd Then
        Window_Detach m_lHostHwnd
    End If
    '/* store hwnd
    m_lHostHwnd = p_OParentObj.hwnd
    m_lHostDc = p_OParentObj.hdc
    
    '/* test for menus
    If GetMenu(m_lHostHwnd) = 0 Then
        m_bCustomMenu = False
    ElseIf p_IMenuBarBg Is Nothing Then
        m_bCustomMenu = False
    End If

    '/* caption bar
    GetObject p_ICaption.handle, Len(tBmp), tBmp
    With tBmp
        m_lTopBorderWidth = .bmWidth
        m_lTopBorderHeight = .bmHeight
    End With
    '/* caption buttons
    GetObject p_ICBoxMin.handle, Len(tBmp), tBmp
    With tBmp
        m_lButtonWidth = (.bmWidth / 3)
        m_lButtonHeight = .bmHeight
    End With

    '/* left border
    GetObject p_ICLeft.handle, Len(tBmp), tBmp
    With tBmp
        m_lLeftBorderWidth = .bmWidth
        m_lLeftBorderHeight = .bmHeight
    End With
    '/* right border
    GetObject p_ICRight.handle, Len(tBmp), tBmp
    With tBmp
        m_lRightBorderWidth = .bmWidth
        m_lRightBorderHeight = .bmHeight
    End With
    '/* bottom height/width
    GetObject p_ICBottom.handle, Len(tBmp), tBmp
    With tBmp
        m_lBottomBorderHeight = .bmHeight
        m_lBottomBorderWidth = .bmWidth
    End With
    
    '/* sanity checks
    If m_lCustomStartPos = 0 Or m_lCustomEndPos = 0 Then
        m_bCustomCaption = False
    End If
    
    '/* custom caption text frame
    If p_ICCapFrame Is Nothing Then
        m_bCaptionFrame = False
    End If
    
    If Not m_bCustomCaption Then
        m_lCustomStartPos = 0
        m_lCustomEndPos = 0
    End If
    
    '/* default caption positioning
    If m_lCaptionOffsetY = 0 Then
        m_lCaptionOffsetY = 4
    End If
    '/* get caption text/icon
    If m_lCaptionOffsetX = 0 Then
        m_lCaptionOffsetX = 24
    End If
    
    '/* default form height
    If m_lMinFormHeight = 0 Then
        m_lMinFormHeight = m_lTopBorderHeight * 3
    End If
    
    '/* get caption text/icon
    Get_Caption m_lHostHwnd
    Get_Icon m_lHostHwnd

    '/* get system default sizes for caption and frame
    m_lLeftMetric = GetSystemMetrics(SM_CXFRAME)
    m_lRightMetric = m_lLeftMetric
    m_lCaptionMetric = GetSystemMetrics(SM_CYCAPTION)
    m_lTopMetric = m_lCaptionMetric + GetSystemMetrics(SM_CYFRAME)
    m_lBottomMetric = GetSystemMetrics(SM_CXFRAME)
    
    '/* render processing
    Set m_cRender = New clsRender
    '/* init image classes
    '/* caption bar
    Set m_cCaption = New clsStoreDc
    m_cCaption.CreateFromPicture p_ICaption
    '/* left border
    Set m_cLeftBorder = New clsStoreDc
    m_cLeftBorder.CreateFromPicture p_ICLeft
    '/* right border
    Set m_cRightBorder = New clsStoreDc
    m_cRightBorder.CreateFromPicture p_ICRight
    '/* bottom border
    Set m_cBottomBorder = New clsStoreDc
    m_cBottomBorder.CreateFromPicture p_ICBottom
    '/* minimize buttons
    Set m_cMinimize = New clsStoreDc
    m_cMinimize.CreateFromPicture p_ICBoxMin
    '/* maximize buttons
    Set m_cMaximize = New clsStoreDc
    m_cMaximize.CreateFromPicture p_ICBoxMax
    '/* restore buttons
    Set m_cRestore = New clsStoreDc
    m_cRestore.CreateFromPicture p_ICBoxRst
    '/* close buttons
    Set m_cClose = New clsStoreDc
    m_cClose.CreateFromPicture p_ICBoxCls
    '/* caption text frame
    If m_bCaptionFrame Then
        Set m_cCaptionFrame = New clsStoreDc
        m_cCaptionFrame.CreateFromPicture p_ICCapFrame
    End If
    
    m_bCanClose = p_CanClose
    m_bCanSize = p_CanSize
    
    '/* start subclassing
    Set m_GSubclass = New GXMSubclass
    Window_Attach m_lHostHwnd
    
    '/* use form icon
    If m_bUseFormIcon Then
        Set_Icon m_lHostDc
    End If

    '/* menu configuration
    If m_bCustomMenu Then
        Set m_cMenu = New clsRCMMenu
        '/* menu bar bg
        Set m_cMenuBarBg = New clsStoreDc
        m_cMenuBarBg.CreateFromPicture p_IMenuBarBg
        '/* menu bar background image
        GetObject p_IMenuBarBg.handle, Len(tBmp), tBmp
        With tBmp
            m_lMenuBarBgHeight = .bmHeight
            m_lMenuBarBgWidth = .bmWidth
            m_lTopBorderHeight = m_lTopBorderHeight + .bmHeight
        End With
        With m_cMenu
            If m_bMenuBar Then Set m_cMenu.p_MenuBarPicture = p_IMenuRollover
            If m_bMenuBg Then Set m_cMenu.p_MenuBgPicture = p_IMenuBg
            Set .p_ImageList = m_vImageList
            If Not m_colIcons Is Nothing Then Set .p_IconIndexSet = m_colIcons
            .p_ActiveMenuForeColor = m_lMenuActiveForeColor
            .p_InActiveMenuForeColor = m_lMenuInActiveForeColor
            .p_MenuBackgroundColor = m_lMenuBackgroundColor
            .p_OfficeXpStyle = m_bOfficeXpStyle
            .p_Transparent = m_bMenuTransparent
            .p_TransIndex = m_btMenuTransIdx
            .p_HighlightStyle = m_eMenuButtonStyle
            .p_CustomRollOver = m_bMenuCustRollOver
            .p_RollOverColor = m_lMenuRollOverColor
            .p_MenuOffsetY = (m_lMenuBarBgHeight + 2)
            .p_CaptionHeight = m_cCaption.Height
            If m_StdMenuFont Is Nothing Then
                .p_MenuFont = p_OParentObj.Font
                Set m_fntMenu = .p_MenuFont
            Else
                .p_MenuFont = m_StdMenuFont
                Set m_fntMenu = m_StdMenuFont
            End If
            '/* attach menu
            .Attach m_lHostHwnd, False
            m_bMenuActive = True
        End With
        If Not p_IMenuRollover Is Nothing Then Set p_IMenuRollover = Nothing
        If Not p_IMenuBg Is Nothing Then Set p_IMenuBg = Nothing
    End If

    Window_Show m_lHostHwnd
    '/* send redraw to caption bar (hack)
    SendMessageLong m_lHostHwnd, WM_NCACTIVATE, 3&, 0&
    
On Error GoTo 0
    
End Sub

Private Function Create_Region(ByVal lWndHdc As Long, _
                               ByVal lHwnd As Long)

'/* create a temporary dc and render image into it
'/* then blit image onto window, excluding client area

Dim lOldBmp     As Long
Dim lNewBmp     As Long
Dim lTmpDc      As Long
Dim lDrawDc     As Long
Dim lActive     As Long
Dim tRect       As RECT

    '/* window size
    GetWindowRect lHwnd, tRect
    OffsetRect tRect, -tRect.left, -tRect.top
    '/* create the dc
    lTmpDc = GetWindowDC(lHwnd)
    lDrawDc = CreateCompatibleDC(lTmpDc)
    lNewBmp = CreateCompatibleBitmap(lTmpDc, tRect.Right, tRect.Bottom)
    lOldBmp = SelectObject(lDrawDc, lNewBmp)

    With tRect
        '/* call drawing sub and paint into temp dc
        Draw_Window lDrawDc, lHwnd, .left, .top, .Right, .Bottom
        '/* select an empty region
        SelectClipRgn lTmpDc, 0
        '/* exclude client area
        ExcludeClipRect lTmpDc, m_lLeftBorderWidth, m_lTopBorderHeight, _
            .Right - .left - m_lRightBorderWidth, .Bottom - .top - m_lBottomBorderHeight
        '/* blit the image
        If m_bCustomCaption Then
            m_cRender.Transparent lTmpDc, .left, .top, .Right, .Bottom, lDrawDc, 0, 0, .Right, .Bottom, &HFF00FF, True
        Else
            m_cRender.Blit lTmpDc, .left, .top, .Right, .Bottom, lDrawDc, 0, 0, SRCCOPY
        End If
    End With

    SelectObject lDrawDc, lOldBmp
    DeleteObject lNewBmp
    DeleteDC lDrawDc
    ReleaseDC lHwnd, lTmpDc

End Function

Public Sub Draw_Rollover()

    
End Sub

Private Sub Draw_Window(ByRef lHdc As Long, _
                        ByVal lHwnd As Long, _
                        ByVal lLeft As Long, _
                        ByVal lTop As Long, _
                        ByVal lRight As Long, _
                        ByVal lBottom As Long)

'/* draw the nc caption and frame

Dim lRightWidth     As Long
Dim lCpEnd          As Long
Dim lBtnLeft        As Long
Dim lStyle          As Long
Dim lRw             As Long
Dim lXe             As Long
Dim hFntMenu        As Long
Dim lTHeight        As Long
Dim lTWidth         As Long
Dim lTempWd         As Long
Dim lTempHt         As Long
Dim menu            As Variant
Dim eState          As ECNBUTTONSTATES
Dim tLF             As LOGFONT
Dim tPos            As POINTAPI
Dim tButtons        As RECT
Dim tButtonR        As RECT
Dim tRect           As RECT
Dim tMenuR          As RECT

'On Error Resume Next
    
    '/* get style context
    lStyle = GetWindowLong(lHwnd, (-16))
    m_bMinimise = ((lStyle And &H20000) = &H20000)
    
    '/* get button position offsets
    If m_bCustomButtonPosition Then
        tButtons.left = lRight - lLeft + m_lControlOffsetX
        tButtons.top = m_lControlOffsetY
    Else
        tButtons.left = lCpEnd + lRightWidth - m_lTopBorderHeight + 4
        tButtons.top = 5
    End If
    
    '/* calculate size/position
    tButtons.Bottom = tButtons.top + m_lButtonHeight
    lBtnLeft = tButtons.left
    GetCursorPos tPos
    GetWindowRect m_lHostHwnd, tRect
    With tPos
        .x = .x - tRect.left
        .y = .y - tRect.top
    End With
    
    lTHeight = (lBottom - lTop) - (m_lTopBorderHeight + m_lBottomBorderHeight)
    lTWidth = (lRight - lLeft) - (m_lLeftBorderWidth + m_lRightBorderWidth)
    
    With m_cRender
        '/* top - form title
        If Not m_bCustomCaption Then
            .Stretch lHdc, 0, 0, lRight - lLeft, m_lTopBorderHeight, m_cCaption.hdc, 0, 0, m_cCaption.Width, m_cCaption.Height, SRCCOPY
        '/* graphic title
        Else
            lTempWd = m_cCaption.Width - m_lCustomEndPos
            lTempHt = m_lTopBorderHeight - m_lMenuBarBgHeight
            .Blit lHdc, 0, 0, m_lCustomStartPos, lTempHt, m_cCaption.hdc, 0, 0, SRCCOPY
            .Stretch lHdc, m_lCustomStartPos, 0, (lRight - lLeft) - (m_lCustomStartPos + lTempWd), lTempHt, m_cCaption.hdc, m_lCustomStartPos, 0, lTempWd, m_cCaption.Height, SRCCOPY
            .Blit lHdc, (lRight - lLeft) - (m_cCaption.Width - m_lCustomEndPos), 0, (lRight - lLeft) - m_lCustomEndPos, lTempHt, m_cCaption.hdc, m_lCustomEndPos, 0, SRCCOPY
        End If
    
        '/* use form caption
        If m_bUseFormCaption Then
            Set_Caption lHdc
        End If
        '/* use form icon
        If m_bUseFormIcon Then
            .Transparent lHdc, (m_lLeftBorderWidth + 4), 4, 16, 16, m_lIconDc, 0, 0, 16, 16, vbMagenta
        End If
    
        '/* left top corner
        .Blit lHdc, 0, 0, m_lLeftBorderWidth, m_lTopBorderHeight, m_cLeftBorder.hdc, 0, 0, SRCCOPY
        '/* left bottom corner
        .Blit lHdc, 0, (lBottom - lTop) - m_lBottomBorderHeight, m_lLeftBorderWidth, m_lBottomBorderHeight, m_cLeftBorder.hdc, 0, m_cLeftBorder.Height - m_lBottomBorderHeight, SRCCOPY
        '/* left side
        .Stretch lHdc, 0, m_lTopBorderHeight, m_lLeftBorderWidth, lTHeight, m_cLeftBorder.hdc, 0, m_lTopBorderHeight, m_cLeftBorder.Width, m_cLeftBorder.Height - (m_lTopBorderHeight + m_lBottomBorderHeight), SRCCOPY
        '/* right top corner
        .Blit lHdc, (lRight - lLeft) - m_lRightBorderWidth, 0, m_lRightBorderWidth, m_lTopBorderHeight, m_cRightBorder.hdc, 0, 0, SRCCOPY
        '/* right bottom corner
        .Blit lHdc, (lRight - lLeft) - m_lRightBorderWidth, (lBottom - lTop) - m_lBottomBorderHeight, m_lRightBorderWidth, m_lBottomBorderHeight, m_cRightBorder.hdc, 0, m_cRightBorder.Height - m_lBottomBorderHeight, SRCCOPY
        '/* right side
        .Stretch lHdc, (lRight - lLeft) - m_lRightBorderWidth, m_lTopBorderHeight, m_lRightBorderWidth, lTHeight, m_cRightBorder.hdc, 0, m_lTopBorderHeight, m_cRightBorder.Width, m_cRightBorder.Height - (m_lTopBorderHeight + m_lBottomBorderHeight), SRCCOPY
        '/* bottom
        .Stretch lHdc, m_lLeftBorderWidth, (lBottom - lTop) - m_lBottomBorderHeight, lTWidth, m_lBottomBorderHeight, m_cBottomBorder.hdc, 0, 0, m_cBottomBorder.Width, m_cBottomBorder.Height, SRCCOPY
    End With

    '/* menu /*
    If m_bCustomMenu Then
        '/* menu bar bg
        m_cRender.Stretch lHdc, m_lLeftBorderWidth, (m_lTopBorderHeight - m_lMenuBarBgHeight), (lRight - lLeft) - (m_lLeftBorderWidth + m_lRightBorderWidth), m_lTopBorderHeight, m_cMenuBarBg.hdc, 0, 0, m_cMenuBarBg.Width, m_cMenuBarBg.Height, SRCCOPY
        '/* set up the menu font
        pOLEFontToLogFont m_fntMenu, lHdc, tLF
        menu = GetMenu(lHwnd)
        If TypeName(menu) = "cPopupMenu" Then
            m_cMenu.p_PopupMenuObject = menu
        Else
            m_cMenu.p_MenuHandle = menu
        End If
        '/* offset sizes
        lRw = (lRight - 1)
        lXe = (lRight - lLeft) - lRw + 1
        If lXe < lLeft + lRw Then
            lXe = lLeft + lRw
        End If
        '/* position menu rollovers
        tLF.lfWeight = FW_NORMAL
        hFntMenu = CreateFontIndirect(tLF)
        With tMenuR
            .left = m_lMenuOffsetX
            .top = (m_lTopBorderHeight \ 2) + m_lMenuOffsetY
            .Bottom = m_lTopBorderHeight
            .Right = lRight - m_lMenuOffsetY
            .top = .top
        End With
        '/* render onto extended caption bar
        With tMenuR
            m_cMenu.Render hFntMenu, lHdc, .left, .top, .Right, _
                .Bottom - .top, -(m_lTopBorderHeight - .top) + 2, True, 10
            .Right = .left
        End With
        '/* cleanup
        DeleteObject hFntMenu
        LSet m_tMenuR = tMenuR
    End If
    
    '/* caption buttons
    '/* get close button size/state
    If m_bCanClose Then
        LSet tButtonR = tButtons
        lBtnLeft = lBtnLeft - (m_lButtonWidth + 1)
        tButtonR.left = lBtnLeft
        tButtonR.Right = tButtonR.left + m_lButtonWidth + 1
        LSet m_tBtn(0) = tButtonR
        If m_bMouseDownClose And PtInRect(tButtonR, tPos.x, tPos.y) Then
            eState = ECN_DOWN
        Else
            eState = ECN_UP
        End If
        NCButton_Draw lHdc, 0, eState, tButtonR
    End If
    
    '/* get max button size/state
    If m_bCanSize Then
        LSet tButtonR = tButtons
        lBtnLeft = lBtnLeft - (m_lButtonWidth + 1)
        With tButtonR
            .left = lBtnLeft
            .Right = .left + m_lButtonWidth + 1
        End With
        LSet m_tBtn(1) = tButtonR
        If m_bMouseDownMinimise And PtInRect(tButtonR, tPos.x, tPos.y) Then
            eState = ECN_DOWN
        Else
            eState = ECN_UP
        End If
        NCButton_Draw lHdc, 1, eState, tButtonR
    End If

    '/* get min button size/state
    If m_bMinimise Then
        LSet tButtonR = tButtons
        lBtnLeft = lBtnLeft - (m_lButtonWidth + 1)
        With tButtonR
            .left = lBtnLeft
            .Right = .left + m_lButtonWidth + 1
        End With
        LSet m_tBtn(2) = tButtonR
        If m_bMouseDownMaximise And PtInRect(tButtonR, tPos.x, tPos.y) Then
            eState = ECN_DOWN
        Else
            eState = ECN_UP
        End If
        NCButton_Draw lHdc, 2, eState, tButtonR
    End If
    
On Error GoTo 0

End Sub

Private Function Get_HitCode() As EHITCONSTANTS

Dim tPos    As POINTAPI
Dim tRect   As RECT
Dim tR      As RECT

    GetCursorPos tPos
    GetWindowRect m_lHostHwnd, tRect
    With tPos
        .x = .x - tRect.left
        .y = .y - tRect.top
    End With
    
    LSet tR = m_tMenuR
    tR.Right = tR.Right + tR.Right
    If PtInRect(tR, tPos.x, tPos.y) <> 0 Then
        Get_HitCode = HTMENU
        Exit Function
    End If
    OffsetRect tRect, -tRect.left, -tRect.top
    Get_HitCode = HTCLIENT
    
    With tPos
        If Not (PtInRect(tRect, .x, .y) = 0) Then
            '/* left
            If .x <= m_lLeftMetric Then
                If .y <= m_lBottomMetric Then
                    If m_bCanSize Then
                        Get_HitCode = HTTOPLEFT
                    End If
                ElseIf .y >= tRect.Bottom - m_lBottomMetric Then
                    If m_bCanSize Then
                        Get_HitCode = HTBOTTOMLEFT
                    End If
                Else
                    If m_bCanSize Then
                        Get_HitCode = HTLEFT
                    End If
                End If
            '/* right
            ElseIf .x >= tRect.Right - m_lRightMetric Then
                If .y <= m_lBottomMetric Then
                    If m_bCanSize Then
                        Get_HitCode = HTTOPRIGHT
                    End If
                ElseIf .y >= tRect.Bottom - m_lBottomMetric Then
                    If m_bCanSize Then
                        Get_HitCode = HTBOTTOMRIGHT
                    End If
                Else
                    If m_bCanSize Then
                        Get_HitCode = HTRIGHT
                    End If
                End If
            '/* top
            ElseIf .y <= m_lTopSizingBorder Then
                If m_bCanSize Then
                    Get_HitCode = HTTOP
                End If
            ElseIf .y >= tRect.Bottom - m_lBottomSizingBorder Then
                If m_bCanSize Then
                    Get_HitCode = HTBOTTOM
                End If
            '/* caption
            ElseIf .y <= m_lTopMetric Then
                If .y <= m_lBottomMetric + m_lCaptionMetric Then
                    Get_HitCode = HTCAPTION
                    If .x <= m_lCaptionMetric Then
                        Get_HitCode = HTSYSMENU
                    End If
                ElseIf .y > m_lBottomMetric + m_lCaptionMetric Then
                    Get_HitCode = HTCLIENT
                End If
            End If
        End If
    End With

End Function

Private Sub Image_Change(ByVal lIndex As Long, _
                         ByVal lDstDc As Long, _
                         ByVal lDstWidth As Long, _
                         ByVal lDstHeight As Long, _
                         ByVal lSrcDc As Long, _
                         ByVal lSrcWidth As Long, _
                         ByVal lSrcHeight As Long, _
                         ByRef tRect As RECT)
    
'/* tristate image swap
    
    Select Case lIndex
    Case 1
        lSrcWidth = lSrcWidth + m_lButtonWidth
    Case 2
        lSrcWidth = lSrcWidth + (m_lButtonWidth * 2)
    End Select
    
    With tRect
        If m_bCustomCaption Then
            m_cRender.Transparent lDstDc, .left, .top, lDstWidth, lDstHeight, lSrcDc, lSrcWidth, lSrcHeight, lDstWidth, lDstHeight, &HFF00FF
        Else
            m_cRender.Blit lDstDc, .left, .top, lDstWidth, lDstHeight, lSrcDc, lSrcWidth, lSrcHeight, SRCCOPY
        End If
    End With

End Sub

Private Sub NCButton_Draw(ByVal lHdc As Long, _
                          ByVal iIndex As Long, _
                          ByVal eState As ECNBUTTONSTATES, _
                          ByRef tRect As RECT)

'/* draw caption bar buttons

Dim lYpos   As Long
Dim lXpos   As Long
Dim lStyle  As Long

    m_bBtnChnge = True
    Select Case iIndex
    Case 0
        If m_bCanClose Then
            Image_Change eState, lHdc, m_lButtonWidth, m_lButtonHeight, m_cClose.hdc, lXpos, lYpos, tRect
        End If
    Case 1
        If m_bCanSize Then
            If m_bMaximized Then
                Image_Change eState, lHdc, m_lButtonWidth, m_lButtonHeight, m_cRestore.hdc, lXpos, lYpos, tRect
            Else
                Image_Change eState, lHdc, m_lButtonWidth, m_lButtonHeight, m_cMaximize.hdc, lXpos, lYpos, tRect
            End If
        End If
    Case 2
        If m_bMinimise Then
            Image_Change eState, lHdc, m_lButtonWidth, m_lButtonHeight, m_cMinimize.hdc, lXpos, lYpos, tRect
            m_lLastPosLeft = lXpos + tRect.left
            m_lLastPosTop = lYpos + tRect.top
        End If
    End Select

End Sub

Private Function NCMouse_Down() As Boolean
'/* caption button image swap

Dim lHdc    As Long
Dim tPos    As POINTAPI
Dim tRect   As RECT

    GetCursorPos tPos
    GetWindowRect m_lHostHwnd, tRect
    lHdc = GetWindowDC(m_lHostHwnd)
    
    With tPos
        .x = .x - tRect.left
        .y = .y - tRect.top
    End With
    With tRect
        OffsetRect tRect, -.left, -.top
    End With
    If m_bCanClose Then
        If PtInRect(m_tBtn(0), tPos.x, tPos.y) <> 0 Then
            NCButton_Draw lHdc, 0, ECN_DOWN, m_tBtn(0)
            m_bMouseDownClose = True
            NCMouse_Down = True
        End If
    End If
    If m_bMaximized Then
        If PtInRect(m_tBtn(1), tPos.x, tPos.y) <> 0 Then
            NCButton_Draw lHdc, 1, ECN_DOWN, m_tBtn(1)
            m_bMouseDownMaximise = True
            NCMouse_Down = True
        End If
    Else
        If PtInRect(m_tBtn(1), tPos.x, tPos.y) <> 0 Then
            NCButton_Draw lHdc, 1, ECN_DOWN, m_tBtn(1)
            m_bMouseDownMaximise = False
            NCMouse_Down = True
        End If
    End If
    If m_bMinimise Then
        If PtInRect(m_tBtn(2), tPos.x, tPos.y) <> 0 Then
            NCButton_Draw lHdc, 2, ECN_DOWN, m_tBtn(2)
            m_bMouseDownMinimise = True
            NCMouse_Down = True
        End If
    End If
    
    ReleaseDC m_lHostHwnd, lHdc

End Function

Private Function NCMouse_Over() As Boolean
'/* caption button image swap

Dim lHdc            As Long
Static sPrevStat    As Long
Dim lCurrStat       As Long
Dim tPos            As POINTAPI
Dim tRect           As RECT

    GetCursorPos tPos
    GetWindowRect m_lHostHwnd, tRect
    lHdc = GetWindowDC(m_lHostHwnd)
    
    '/* get cursor pos and offset
    With tPos
        .x = .x - tRect.left
        .y = .y - tRect.top
        If Not PtInRect(m_tBtn(0), .x, .y) = 0 Then
            '/* swap image
            If Not sPrevStat = 1 Then
                NCButton_Draw lHdc, 0, ECN_OVER, m_tBtn(0)
            End If
            lCurrStat = 1
        ElseIf Not PtInRect(m_tBtn(1), .x, .y) = 0 Then
            If Not sPrevStat = 2 Then
                NCButton_Draw lHdc, 1, ECN_OVER, m_tBtn(1)
            End If
            lCurrStat = 2
        ElseIf Not PtInRect(m_tBtn(2), .x, .y) = 0 Then
            If Not sPrevStat = 3 Then
                NCButton_Draw lHdc, 2, ECN_OVER, m_tBtn(2)
            End If
            lCurrStat = 3
        Else
            If sPrevStat > 0 Then
                NCButton_Draw lHdc, (sPrevStat - 1), ECN_UP, m_tBtn(sPrevStat - 1)
            End If
        End If
        '/* conditional reset
        If lCurrStat <> sPrevStat Then
            If sPrevStat > 0 Then
                NCButton_Draw lHdc, (sPrevStat - 1), ECN_UP, m_tBtn(sPrevStat - 1)
            End If
        End If
        sPrevStat = lCurrStat
    End With
    
    ReleaseDC m_lHostHwnd, lHdc

End Function

Private Sub NCMouse_Up()
'/* caption button events

Dim lStyle  As Long
Dim lHdc    As Long
Dim tPos    As POINTAPI
Dim tRect   As RECT

    GetCursorPos tPos
    GetWindowRect m_lHostHwnd, tRect
    lHdc = GetWindowDC(m_lHostHwnd)
    
    With tPos
        .x = .x - tRect.left
        .y = .y - tRect.top
    End With
    With tRect
        OffsetRect tRect, -.left, -.top
    End With
    '/* close window
    If m_bCanClose Then
        If Not PtInRect(m_tBtn(0), tPos.x, tPos.y) = 0 Then
            If m_bMouseDownClose Then
                Window_Fade m_lHostHwnd
                Window_Post SC_CLOSE
            End If
        End If
    End If
    '/* maximize
    If Not PtInRect(m_tBtn(1), tPos.x, tPos.y) = 0 Then
        If m_bMaximized Then
            Window_Restore
        Else
            Window_Maximize
        End If
    End If
    '/* minimize
    If m_bMinimise Then
        If Not PtInRect(m_tBtn(2), tPos.x, tPos.y) = 0 Then
            If m_bMouseDownMinimise Then
                Window_Post SC_MINIMIZE
            End If
        End If
    End If
    
    ReleaseDC m_lHostHwnd, lHdc
    Reset_Caption

End Sub

Private Sub pOLEFontToLogFont(fntThis As StdFont, _
                              ByVal lHdc As Long, _
                              tLF As LOGFONT)

'/* Convert an OLE StdFont to a LOGFONT structure

Dim sFont   As String

On Error Resume Next

    With tLF
        sFont = fntThis.Name
        .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(lHdc, LOGPIXELSY)), 72)
        .lfItalic = fntThis.Italic
        If fntThis.Bold Then
            .lfWeight = FW_BOLD
        Else
            .lfWeight = FW_NORMAL
        End If
        .lfUnderline = fntThis.Underline
        .lfStrikeOut = fntThis.Strikethrough
        .lfCharSet = fntThis.Charset
    End With

On Error GoTo 0

End Sub

Private Function TranslateColor(ByVal Clr As OLE_COLOR, _
                                Optional hPal As Long = 0) As Long

    If OleTranslateColor(Clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If

End Function

Private Sub Reset_Caption()
'/* reset caption bar images

Dim lHdc    As Long

    lHdc = GetWindowDC(m_lHostHwnd)
    NCButton_Draw lHdc, 0, ECN_UP, m_tBtn(0)
    NCButton_Draw lHdc, 1, ECN_UP, m_tBtn(1)
    NCButton_Draw lHdc, 2, ECN_UP, m_tBtn(2)
    m_bMouseDownMinimise = False
    m_bMouseDownMaximise = False
    m_bMouseDownClose = False
    m_bBtnChnge = False
    ReleaseDC m_lHostHwnd, lHdc

End Sub

Private Sub Get_Caption(ByVal lHwnd As Long)

'/* get caption text from hosted form
Dim tPnt    As POINTAPI

'On Error Resume Next

    m_sCapText = p_OParentObj.Caption
    If Len(m_sCapText) = 0 Then
        m_sCapText = "NeoClass GII"
    End If
    
    '/* get text height/width
    GetTextExtentPoint32 m_lHostDc, m_sCapText, Len(m_sCapText) + 2, tPnt
    '/* borrowed from Paul Catons example
    '/* set minimum width
    '/* via caption/button size
    With m_SizeInfo
        '/* minimum size
        With .ptMinTrackSize
            .x = (((m_lButtonWidth + 1) * 3) - m_lControlOffsetX) + tPnt.x + (m_lCaptionOffsetX + 24)
            If .x < m_lMinFormWidth Then
                .x = m_lMinFormWidth
            End If
            .y = m_lMinFormHeight
        End With
        '/* max pos
        With .ptMaxPosition
            .x = 0
            .y = 0
        End With
        '/* Maximized size
        With .ptMaxSize
            .x = Screen.Width
            .y = Screen.Height
        End With
        '/* Maximum size with re-sizing
        With .ptMaxTrackSize
            .x = Screen.Width
            .y = Screen.Height
        End With
    End With

On Error GoTo 0

End Sub

Private Sub Set_Caption(lHdc As Long)
'/* draw caption text

Dim lFontOld    As Long
Dim lHFont      As Long
Dim tLogFont    As LOGFONT
Dim tPnt        As POINTAPI
Dim tRect       As RECT
Dim tWndR       As RECT

'On Error Resume Next

    '/* create font structure based on main
    '/* form settings (inheritence provided)
    With tLogFont
        .lfFaceName = p_OParentObj.Font + vbNullChar
        .lfHeight = (p_OParentObj.Font.Size * -20) / Screen.TwipsPerPixelY
        .lfWeight = 700
        .lfItalic = IIf(p_OParentObj.FontItalic, 1, 0)
        .lfUnderline = IIf(p_OParentObj.FontUnderline, 1, 0)
        .lfStrikeOut = IIf(p_OParentObj.FontStrikethru, 1, 0)
        .lfQuality = ANTIALIASED_QUALITY
    End With
    
    '/* get text height/width
    GetTextExtentPoint32 lHdc, m_sCapText, lstrlen(m_sCapText) + 2, tPnt
    '/* create font
    lHFont = CreateFontIndirect(tLogFont)
    '/* place in dc
    lFontOld = SelectObject(lHdc, lHFont)
    '/* bg transparent
    SetBkMode lHdc, 1
    '/* set font color
    SetTextColor lHdc, m_CaptionFntClr
    
    '/* exclusive properties
    If m_bCaptionFrame = True Then
        m_bCenterCaption = False
    End If
    
    '/* draw text
    With tRect
        .top = m_lCaptionOffsetY
        .Bottom = .top + tPnt.y
        If m_bCenterCaption Then
            '/* thanks enmity
            GetWindowRect m_lHostHwnd, tWndR
            .left = ((tWndR.Right - tWndR.left) - tPnt.x) / 2 - 10
            .Right = .left + tPnt.x + 10
        Else
            .left = (m_lLeftBorderWidth + m_lCaptionOffsetX)
            .Right = tPnt.x + m_lCaptionOffsetX
        End If
    End With
    
    If m_bCaptionFrame Then
        With tRect
            m_cRender.Stretch lHdc, tRect.left, 1, .Right, m_cCaptionFrame.Height, m_cCaptionFrame.hdc, 0, 0, m_cCaptionFrame.Width, m_cCaptionFrame.Height, SRCCOPY, True
            .left = .left + 10
            .Right = .Right + 10
            .top = .top + 2
        End With
        DrawText lHdc, m_sCapText, -1, tRect, &H1 Or &H20 Or &H4
    Else
        DrawText lHdc, m_sCapText, -1, tRect, &H1 Or &H20 Or &H4
    End If

    '/* delete fonts
    SelectObject lHdc, lFontOld
    DeleteObject lHFont

On Error GoTo 0

End Sub

Public Sub Change_Caption(ByVal sCaption As String)

'On Error Resume Next

    If Not sCaption = m_sCapText Then
        m_sCapText = sCaption
    End If
    Window_Paint m_lHostHwnd

On Error GoTo 0

End Sub

Private Sub Get_Icon(ByVal lHwnd As Long)
'/* get form icon handle
    m_lCapIcon = p_OParentObj.Icon.handle
End Sub

Private Sub Set_Icon(ByVal lHdc As Long)
'/* draw icon

Dim lSDc    As Long
Dim lBmp    As Long
Dim lTop    As Long
Dim lBmpOld As Long
Dim lLeft   As Long
Dim tRect   As RECT
Dim lLBrsh  As LOGBRUSH

On Error Resume Next

    '/* create a temp image dc
    m_lIconDc = CreateCompatibleDC(m_lHostDc)
    m_lIconBmp = CreateCompatibleBitmap(m_lHostDc, 16, 16)
    m_lIconBmpOld = SelectObject(m_lIconDc, m_lIconBmp)
    With tRect
        .Right = 16
        .Bottom = 16
    End With
    '/* backpaint
    lLBrsh.lbColor = vbMagenta
    lBmp = CreateBrushIndirect(lLBrsh)
    FillRect m_lIconDc, tRect, lBmp
    DeleteObject lBmp
    '/* draw icon
    DrawIconEx m_lIconDc, 0&, 0&, m_lCapIcon, 16&, 16&, 0&, 0&, DI_NORMAL
    
On Error GoTo 0

End Sub

Private Sub Destroy_Icon()

    SelectObject m_lIconDc, m_lIconBmpOld
    DeleteObject m_lIconBmp
    DeleteDC m_lIconDc
    m_lIconBmpOld = 0
    m_lIconBmp = 0
    m_lIconDc = 0
    
End Sub

Private Sub Window_Fade(ByVal lHwnd As Long)

Dim lStyle  As Long
Dim lFdIdx  As Long
Dim lCt     As Long

    If Not m_bWin32 Then Exit Sub
    If m_bFadeOutEffect Then
        lStyle = GetWindowLong(lHwnd, GWL_EXSTYLE)
        SetWindowLong lHwnd, GWL_EXSTYLE, lStyle Or WS_EX_LAYERED
        lFdIdx = 255
        For lCt = 255 To 10 Step -2
            SetLayeredWindowAttributes lHwnd, 0&, lCt, LWA_ALPHA
            DoEvents
        Next lCt
    End If

End Sub

Private Sub Window_Move()
'/* respond to window movement

Dim lParent     As Long
Dim lXpos       As Long
Dim lYpos       As Long
Dim tPInit      As POINTAPI
Dim tPLast      As POINTAPI
Dim tPos        As POINTAPI
Dim tRect       As RECT

    '/* if not maximized get size
    If Not m_bMaximized Then
        GetWindowRect m_lHostHwnd, tRect
        lParent = GetParent(m_lHostHwnd)
        '/* map position
        If Not lParent = 0 Then
            MapWindowPoints 0&, lParent, tRect, 2
        End If
        GetCursorPos tPInit
        LSet tPLast = tPInit
        '/* map pos and move window
        Do While Not (GetAsyncKeyState(vbLeftButton) = 0)
            GetCursorPos tPos
            With tPos
                If Not .x = tPLast.x Or Not .y = tPLast.y Then
                    lXpos = .x - tPLast.x
                    lYpos = .y - tPLast.y
                    With tRect
                        SetWindowPos m_lHostHwnd, 0&, .left + lXpos, .top + lYpos, _
                            0&, 0&, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOOWNERZORDER
                    End With
                    LSet tPLast = tPos
                    GetWindowRect m_lHostHwnd, tRect
                    If Not lParent = 0 Then
                        MapWindowPoints 0&, lParent, tRect, 2
                    End If
                End If
            End With
            DoEvents
            Sleep 1
        Loop
    End If
    Window_Redraw m_lHostHwnd

End Sub

Private Sub Window_Minimize()
    Window_Store
End Sub

Private Sub Window_Maximize()

Dim lX      As Long
Dim lY      As Long
Dim lW      As Long
Dim lH      As Long
Dim lCt     As Long
Dim tMax    As RECT
Dim tTemp   As RECT

    Window_Store
    Window_Metrics WD_MAXIMIZED, tMax
    
    m_bMaximized = True
    With tMax
        lW = .Right / 8
        lH = .Bottom / 8
    End With
    
    LSet tTemp = m_tSizeState
    With tTemp
        OffsetRect tTemp, -.left, -.top
        .left = m_tSizeState.left
        .top = m_tSizeState.top
    End With
    
    With tTemp
        lX = .left / 8
        lY = .top / 8
        For lCt = 1 To 9
            .top = .top - lY
            If .top < tMax.top Then .top = tMax.top
            .left = .left - lX
            If .left < tMax.left Then .left = tMax.left
            .Right = .Right + lW
            If .Right > tMax.Right Then .Right = tMax.Right
            .Bottom = .Bottom + lH
            If .Bottom > tMax.Bottom Then .Bottom = tMax.Bottom
            MoveWindow m_lHostHwnd, .left, .top, .Right, .Bottom, 1
            DoEvents
            Sleep 5
        Next lCt
        '/* if window is offscreen
        MoveWindow m_lHostHwnd, tMax.left, tMax.top, tMax.Right, tMax.Bottom, 1
    End With
    
End Sub

Private Sub Window_Restore()

Dim lX      As Long
Dim lY      As Long
Dim lW      As Long
Dim lH      As Long
Dim lCt     As Long
Dim tMax    As RECT
Dim tTemp   As RECT

    Window_Metrics WD_MAXIMIZED, tMax
    
    With m_tSizeState
        lX = .left / 8
        lY = .top / 8
    End With
    
    LSet tTemp = m_tSizeState
    With tTemp
        OffsetRect tTemp, -.left, -.top
    End With
    
    With tMax
        lW = .Right / 8
        lH = .Bottom / 8
        For lCt = 9 To 1 Step -1
            .top = .top + lY
            If .top > m_tSizeState.top Then .top = m_tSizeState.top
            .left = .left + lX
            If .left > m_tSizeState.left Then .left = m_tSizeState.left
            .Right = .Right - lW
            If .Right < tTemp.Right Then .Right = tTemp.Right
            .Bottom = .Bottom - lH
            If .Bottom < tTemp.Bottom Then .Bottom = tTemp.Bottom
            MoveWindow m_lHostHwnd, .left, .top, .Right, .Bottom, 1
            DoEvents
            Sleep 5
        Next lCt
        MoveWindow m_lHostHwnd, m_tSizeState.left, m_tSizeState.top, tTemp.Right, tTemp.Bottom, 1
    End With

    m_bMaximized = False

End Sub

Private Sub Window_Store()
    '/* store current size
    GetWindowRect m_lHostHwnd, m_tSizeState
End Sub

Private Sub Window_Metrics(ByRef eSizeState As WD_STATE, _
                           ByRef tRect As RECT)

    Select Case eSizeState
    Case WD_MINIMIZED
        With tRect
            .Bottom = m_lMinFormHeight
            .left = 0
            .Right = m_lMinFormWidth
            .top = 0
        End With
    
    Case WD_MAXIMIZED
        Dim lHDiff, lWDiff As Long
        lHDiff = (m_lCaptionMetric + m_lBottomMetric) - (m_lTopBorderHeight + (m_lBottomBorderHeight * 2))
        lWDiff = (m_lLeftMetric * 2) - (m_lLeftBorderWidth + m_lRightBorderWidth)
        With tRect
            .Bottom = (GetSystemMetrics(SM_CYFULLSCREEN) - lHDiff) + 4
            .left = (lWDiff / 2)
            .Right = GetSystemMetrics(SM_CXFULLSCREEN) - lWDiff
            .top = -4
        End With
    
    Case WD_NORMAL
        LSet tRect = m_tSizeState
    End Select

End Sub

Private Sub Window_Show(ByVal lHwnd As Long)
'/* show window
    SetWindowPos lHwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED Or SWP_NOOWNERZORDER
End Sub

Private Sub Window_Post(ByVal eCmd As ECNCMD)
'/* post change to message queue
    PostMessage m_lHostHwnd, &H112, eCmd, 0
End Sub

Private Sub Window_Redraw(lHwnd As Long)
'/* forced repaint
    RedrawWindow lHwnd, ByVal 0&, 0&, RDW_NRPGROUP
End Sub

Private Sub Window_Paint(ByVal lHwnd As Long)

Dim lHdc As Long

    lHdc = GetWindowDC(lHwnd)
    Create_Region lHdc, lHwnd
    ReleaseDC lHwnd, lHdc
        
End Sub

Private Sub Window_Attach(ByVal lHwnd As Long)
'/* attatch messages
    
Dim lThread As Long

    If Not lHwnd = 0 Then
        With m_GSubclass
            .Subclass lHwnd, Me
            .AddMessage lHwnd, WM_ACTIVATE, MSG_BEFORE
            .AddMessage lHwnd, WM_NCACTIVATE, MSG_BEFORE
            .AddMessage lHwnd, WM_SETCURSOR, MSG_BEFORE
            .AddMessage lHwnd, WM_GETMINMAXINFO, MSG_BEFORE
            .AddMessage lHwnd, WM_NCCALCSIZE, MSG_BEFORE
            .AddMessage lHwnd, WM_NCHITTEST, MSG_BEFORE
            .AddMessage lHwnd, WM_NCPAINT, MSG_BEFORE
            .AddMessage lHwnd, WM_INITMENU, MSG_BEFORE
            .AddMessage lHwnd, WM_MOUSEMOVE, MSG_BEFORE
            .AddMessage lHwnd, WM_NCMOUSEMOVE, MSG_BEFORE
            .AddMessage lHwnd, WM_NCLBUTTONDOWN, MSG_BEFORE
            .AddMessage lHwnd, WM_NCLBUTTONUP, MSG_BEFORE
            .AddMessage lHwnd, WM_NCLBUTTONDBLCLK, MSG_BEFORE
            If m_bCustomMenu Then
                .AddMessage lHwnd, WM_INITMENUPOPUP, MSG_BEFORE
                .AddMessage lHwnd, WM_ENTERMENULOOP, MSG_BEFORE
                .AddMessage lHwnd, WM_EXITMENULOOP, MSG_BEFORE
                lThread = GetCurrentThreadId
                If Not lThread <= 0 Then InstallKeyboardHook Me
            End If
        End With
    End If

End Sub

Private Sub Window_Detach(ByVal lHwnd As Long)
'/* detatch messages

Dim lMenu   As Long

    If lHwnd = 0 Then Exit Sub
    With m_GSubclass
        .DeleteMessage lHwnd, WM_ACTIVATE, MSG_BEFORE
        .DeleteMessage lHwnd, WM_NCACTIVATE, MSG_BEFORE
        .DeleteMessage lHwnd, WM_SETCURSOR, MSG_BEFORE
        .DeleteMessage lHwnd, WM_GETMINMAXINFO, MSG_BEFORE
        .DeleteMessage lHwnd, WM_NCCALCSIZE, MSG_BEFORE
        .DeleteMessage lHwnd, WM_NCHITTEST, MSG_BEFORE
        .DeleteMessage lHwnd, WM_NCPAINT, MSG_BEFORE
        .DeleteMessage lHwnd, WM_INITMENU, MSG_BEFORE
        .DeleteMessage lHwnd, WM_MOUSEMOVE, MSG_BEFORE
        .DeleteMessage lHwnd, WM_NCMOUSEMOVE, MSG_BEFORE
        .DeleteMessage lHwnd, WM_NCLBUTTONDOWN, MSG_BEFORE
        .DeleteMessage lHwnd, WM_NCLBUTTONUP, MSG_BEFORE
        .DeleteMessage lHwnd, WM_NCLBUTTONDBLCLK, MSG_BEFORE
        If m_bCustomMenu Then
            .DeleteMessage lHwnd, WM_INITMENUPOPUP, MSG_BEFORE
            .DeleteMessage lHwnd, WM_ENTERMENULOOP, MSG_BEFORE
            .DeleteMessage lHwnd, WM_EXITMENULOOP, MSG_BEFORE
        End If
        .UnSubclass lHwnd
    End With
    
End Sub

Private Sub GXISubclass_WndProc(ByVal bBefore As Boolean, _
                                bHandled As Boolean, _
                                lReturn As Long, _
                                ByVal lHwnd As Long, _
                                ByVal uMsg As WinSubHookGX.eMsg, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long, _
                                lParamUser As Long)

Dim lStyle          As Long
Dim lNewStyle       As Long
Dim lActWnd         As Long
Dim lHdc            As Long
Dim eHitCode        As EHITCONSTANTS
Dim tRect           As RECT
Dim tWP             As WINDOWPOS
Dim tNCR            As NCCALCSIZE_PARAMS

    Select Case uMsg
    '/* init menu popup
    Case WM_INITMENUPOPUP
        If Not (lParam And &HFFFF0000) = 0 Then
            '/* Sys Menu:
            m_cMenu.OwnerDrawMenu wParam
        Else
            '/* App Menu:
            m_cMenu.OwnerDrawMenu wParam
        End If
        Window_Paint lHwnd
        bHandled = True

    '/* menu init
    Case WM_INITMENU
        If Not Get_HitCode = 1 Then
            Window_Paint lHwnd
        End If
        bHandled = True
        
    '/* enter and exit menu loop
    Case WM_ENTERMENULOOP, WM_EXITMENULOOP
        If uMsg = WM_EXITMENULOOP Then
            lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        Else
            If wParam = 0 Then
                lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
                Window_Paint lHwnd
            End If
        End If
        bHandled = True

    '/* active
    Case WM_ACTIVATE
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        If lHwnd = GetActiveWindow Then
            '/* force a calc message to redraw header
            If Not GetAsyncKeyState(vbLeftButton) = 0 Then
                Window_Paint lHwnd
                Window_Redraw lHwnd
            End If
        End If
        bHandled = True
    
    '/* remove header -some nifty hacks! if we let windows have this
    '/* it will flicker when a new form loads overtop
    Case WM_NCACTIVATE
        '/* first run
        If wParam = 3 Then
            Window_Paint lHwnd
            Window_Redraw lHwnd
        End If
        If Not GetAsyncKeyState(vbLeftButton) = 0 Then
            lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        End If
        Window_Paint lHwnd
        bHandled = True
        
    '/* cursor
    Case WM_SETCURSOR
        lStyle = GetWindowLong(m_lHostHwnd, GWL_STYLE)
        lNewStyle = lStyle And Not &H10000000
        SetWindowLong m_lHostHwnd, GWL_STYLE, lNewStyle
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        SetWindowLong lHwnd, GWL_STYLE, lStyle
        bHandled = True
        
    '/* min size
    Case WM_GETMINMAXINFO
        CopyMemory ByVal lParam, m_SizeInfo, LenB(m_SizeInfo)
        bHandled = True

    '/* calculate size of frame
    Case WM_NCCALCSIZE
        If Not wParam = 0 Then
            CopyMemory tNCR, ByVal lParam, Len(tNCR)
            CopyMemory tWP, ByVal tNCR.lppos, Len(tWP)
            '/* get new size
            With tNCR.rgrc(0)
                .left = tWP.x
                .top = tWP.y
                .Right = tWP.x + tWP.cx
                .Bottom = tWP.y + tWP.cy
                '/* set size of new rectangle
                .left = .left + m_lLeftBorderWidth
                .top = .top + m_lTopBorderHeight
                .Right = .Right - m_lRightBorderWidth
                .Bottom = .Bottom - m_lBottomBorderHeight
            End With
            LSet tNCR.rgrc(1) = tNCR.rgrc(0)
            CopyMemory ByVal lParam, tNCR, Len(tNCR)
            lReturn = &H400
        Else
            lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        End If
        bHandled = True

    '/* hit test
    Case WM_NCHITTEST
        eHitCode = Get_HitCode()
        m_eLastHT = eHitCode
        If eHitCode = HTMENU Then
             lReturn = HTCLIENT
        Else
            lReturn = eHitCode
        End If
        bHandled = True

    '/* paint
    Case WM_NCPAINT
        If IsWindowVisible(lHwnd) = 0 Then
            lReturn = 0
            bHandled = True
            Exit Sub
        End If
        Window_Paint lHwnd
        bHandled = True
        
    '/* form mouse move
    Case WM_MOUSEMOVE
        If m_bBtnChnge Then
            Reset_Caption
        End If
    
    '/* button mouse over
    Case WM_NCMOUSEMOVE
        NCMouse_Over
        bHandled = True

    '/* nc mouse down
    Case WM_NCLBUTTONDOWN
        If NCMouse_Down Then
            bHandled = True
            Exit Sub
        End If
        wParam = Get_HitCode()
        If m_eLastHT = HTCAPTION Then
            If Not m_bMaximized Then
                Window_Move
                bHandled = True
            Else
                lReturn = 1
                bHandled = True
            End If
            Exit Sub
        End If
        lReturn = m_GSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        Window_Paint lHwnd
        bHandled = True
    
    '/* button up
    Case WM_NCLBUTTONUP
        NCMouse_Down
        NCMouse_Up
        Window_Paint lHwnd
        bHandled = True
        
    '/* caption dbl click
    Case WM_NCLBUTTONDBLCLK
        If Not m_bCanSize Then
            Exit Sub
        End If
        If m_bMaximized Then
            Window_Restore
        Else
            Window_Maximize
        End If
        bHandled = True
    End Select

End Sub

Public Sub Clean_Up()
'/* purge resources

On Error GoTo Handler

    If Not m_lHostHwnd = 0 Then
        '/* reset the thread
        If m_bGThreadAccl Then
            Dim lThread As Long
            lThread = GetCurrentThreadId
            If Not lThread <= 0 Then
                mControl.ThreadAccelerate lThread, Thread_Normal
            End If
        End If
        '/* detatch the hook!
        RemoveKeyboardHook
        Window_Detach m_lHostHwnd
        '/* destroy classes
        If Not m_GSubclass Is Nothing Then Set m_GSubclass = Nothing
        If Not p_OParentObj Is Nothing Then Set p_OParentObj = Nothing
        If Not m_cMenu Is Nothing Then Set m_cMenu = Nothing
        If Not m_cRender Is Nothing Then Set m_cRender = Nothing
        If Not m_cCaption Is Nothing Then Set m_cCaption = Nothing
        If Not m_cLeftBorder Is Nothing Then Set m_cLeftBorder = Nothing
        If Not m_cRightBorder Is Nothing Then Set m_cRightBorder = Nothing
        If Not m_cBottomBorder Is Nothing Then Set m_cBottomBorder = Nothing
        If Not m_cMinimize Is Nothing Then Set m_cMinimize = Nothing
        If Not m_cMaximize Is Nothing Then Set m_cMaximize = Nothing
        If Not m_cClose Is Nothing Then Set m_cClose = Nothing
        If Not m_cRestore Is Nothing Then Set m_cRestore = Nothing
        If Not m_cMenuBarBg Is Nothing Then Set m_cMenuBarBg = Nothing
        If Not m_cCaptionFrame Is Nothing Then Set m_cCaptionFrame = Nothing
        '/* unload objects
        If Not m_StdMenuFont Is Nothing Then Set m_StdMenuFont = Nothing
        If Not m_colIcons Is Nothing Then Set m_colIcons = Nothing
        Destroy_Icon
        If Not p_ICaption Is Nothing Then Set p_ICaption = Nothing
        If Not p_ICBottom Is Nothing Then Set p_ICBottom = Nothing
        If Not p_ICLeft Is Nothing Then Set p_ICLeft = Nothing
        If Not p_ICRight Is Nothing Then Set p_ICRight = Nothing
        If Not p_ICBoxMin Is Nothing Then Set p_ICBoxMin = Nothing
        If Not p_ICBoxMax Is Nothing Then Set p_ICBoxMax = Nothing
        If Not p_ICBoxRst Is Nothing Then Set p_ICBoxRst = Nothing
        If Not p_ICBoxCls Is Nothing Then Set p_ICBoxCls = Nothing
        If Not p_IMenuBg Is Nothing Then Set p_IMenuBg = Nothing
        If Not p_IMenuRollover Is Nothing Then Set p_IMenuRollover = Nothing
        If Not p_IMenuBarBg Is Nothing Then Set p_IMenuBarBg = Nothing
    End If
    m_lHostHwnd = 0

Exit Sub

Handler:
    Error_Data "clsrcm", "clean_up", Err.Number, Err.Description
    On Error GoTo 0

End Sub

Private Sub Class_Terminate()
    Clean_Up
End Sub
