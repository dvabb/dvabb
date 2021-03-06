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
 * File Name: fcktoolbarspecialcombo.js
 * 	FCKToolbarSpecialCombo Class: This is a "abstract" base class to be used
 * 	by the special combo toolbar elements like font name, font size, paragraph format, etc...
 * 	
 * 	The following properties and methods must be implemented when inheriting from
 * 	this class:
 * 		- Property:	Command								[ The command to be executed ]
 * 		- Method:	GetLabel()							[ Returns the label ]
 * 		-			CreateItems( targetSpecialCombo )	[ Add all items in the special combo ]
 * 
 * Version:  2.0 RC3
 * Modified: 2005-01-04 18:41:03
 * 
 * File Authors:
 * 		Frederico Caldeira Knabben (fredck@fckeditor.net)
 */

var FCKToolbarSpecialCombo = function()
{
	this.SourceView			= false ;
	this.ContextSensitive	= true ;
}

FCKToolbarSpecialCombo.prototype.CreateInstance = function( parentToolbar )
{
	this._Combo = new FCKSpecialCombo( this.GetLabel() ) ;
	this._Combo.FieldWidth = 100 ;
	this._Combo.PanelWidth = 150 ;
	this._Combo.PanelMaxHeight = 150 ;
	
	this.CreateItems( this._Combo ) ;

	this._Combo.Create( parentToolbar.DOMRow.insertCell(-1) ) ;

	this._Combo.Command = this.Command ;
	
	this._Combo.OnSelect = function( itemId, item )
	{
		this.Command.Execute( itemId, item ) ;
	}
}

FCKToolbarSpecialCombo.prototype.RefreshState = function()
{
	// Gets the actual state.
	var eState ;
	
//	if ( FCK.EditMode == FCK_EDITMODE_SOURCE && ! this.SourceView )
//		eState = FCK_TRISTATE_DISABLED ;
//	else
//	{
		var sValue = this.Command.GetState() ;

		if ( sValue != FCK_TRISTATE_DISABLED )
		{
			eState = FCK_TRISTATE_ON ;
			
			if ( !this.RefreshActiveItems )
			{
				this.RefreshActiveItems = function( combo, value )
				{
					this._Combo.DeselectAll() ;
					this._Combo.SelectItem( value ) ;
					this._Combo.SetLabelById( value ) ;
				}
			}
			this.RefreshActiveItems( this._Combo, sValue ) ;
		}
		else
			eState = FCK_TRISTATE_DISABLED ;
//	}
	
	// If there are no state changes then do nothing and return.
	if ( eState == this.State ) return ;
	
	if ( eState == FCK_TRISTATE_DISABLED )
	{
		this._Combo.DeselectAll() ;
		this._Combo.SetLabel( '' ) ;
	}

	// Sets the actual state.
	this.State = eState ;

	// Updates the graphical state.
	this._Combo.SetEnabled( eState != FCK_TRISTATE_DISABLED ) ;
}

FCKToolbarSpecialCombo.prototype.Enable = function()
{
	this.RefreshState() ;
}

FCKToolbarSpecialCombo.prototype.Disable = function()
{
	this.State = FCK_TRISTATE_DISABLED ;
	this._Combo.DeselectAll() ;
	this._Combo.SetLabel( '' ) ;
	this._Combo.SetEnabled( false ) ;
}