/*
 * FCKeditor - The text editor for internet
 * Copyright (C) 2003-2004 Frederico Caldeira Knabben
 * 
 * Licensed under the terms of the GNU Lesser General Public License:
 * 		http://www.opensource.org/licenses/lgpl-license.php
 * 
 * For further information visit:
 * 		http://www.fckeditor.net/
 * 
 * File Name: fcktoolbarbreak_ie.js
 * 	FCKToolbarBreak Class: breaks the toolbars.
 * 	It makes it possible to force the toolbar to brak to a new line.
 * 	This is the IE specific implementation.
 * 
 * Version:  2.0 RC3
 * Modified: 2005-02-09 18:07:44
 * 
 * File Authors:
 * 		Frederico Caldeira Knabben (fredck@fckeditor.net)
 */

var FCKToolbarBreak = function()
{
	var oBreakDiv = document.createElement( 'div' ) ;
	
	oBreakDiv.className = 'TB_Break' ;
	
	oBreakDiv.style.clear = FCKLang.Dir == 'rtl' ? 'left' : 'right' ;
	
	FCKToolbarSet.DOMElement.appendChild( oBreakDiv ) ;
}