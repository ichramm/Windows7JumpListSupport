/*
 * Copyright (C) 2011 by Juan Ramirez (@ichramm) - jramirez@inconcertcc.com
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */

/*
Do not use #pragma once, as this file needs to be included twice.  Once to declare the externs
for the GUIDs, and again right after including initguid.h to actually define the GUIDs.
*/



// package guid
// { a82752e2-f1e0-48ff-adeb-6e7ca1626c21 }
#define guidWindows7JumpListSupportPkg { 0xA82752E2, 0xF1E0, 0x48FF, { 0xAD, 0xEB, 0x6E, 0x7C, 0xA1, 0x62, 0x6C, 0x21 } }
#ifdef DEFINE_GUID
DEFINE_GUID(CLSID_Windows7JumpListSupport,
0xA82752E2, 0xF1E0, 0x48FF, 0xAD, 0xEB, 0x6E, 0x7C, 0xA1, 0x62, 0x6C, 0x21 );
#endif

// Command set guid for our commands (used with IOleCommandTarget)
// { 9f496474-0ca8-4771-b57b-73f59e4a4004 }
#define guidWindows7JumpListSupportCmdSet { 0x9F496474, 0xCA8, 0x4771, { 0xB5, 0x7B, 0x73, 0xF5, 0x9E, 0x4A, 0x40, 0x4 } }
#ifdef DEFINE_GUID
DEFINE_GUID(CLSID_Windows7JumpListSupportCmdSet, 
0x9F496474, 0xCA8, 0x4771, 0xB5, 0x7B, 0x73, 0xF5, 0x9E, 0x4A, 0x40, 0x4 );
#endif

//Guid for the image list referenced in the VSCT file
// { a253d11e-b143-4b60-98d1-4f0046cb65db }
#define guidImages { 0xA253D11E, 0xB143, 0x4B60, { 0x98, 0xD1, 0x4F, 0x0, 0x46, 0xCB, 0x65, 0xDB } }
#ifdef DEFINE_GUID
DEFINE_GUID(CLSID_Images, 
0xA253D11E, 0xB143, 0x4B60, 0x98, 0xD1, 0x4F, 0x0, 0x46, 0xCB, 0x65, 0xDB );
#endif


