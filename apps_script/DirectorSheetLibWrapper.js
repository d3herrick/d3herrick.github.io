//
// Copyright 2023 Douglas Herrick
//
// Use of this source code is governed by an MIT-style
// license that can be found in the LICENSE file or at
//
// https://opensource.org/licenses/MIT.
//
// This collection of wrapper functions belay the interesting work to the library that manages
// the spreadsheets that Newton Tree Conservancy directors use to manage planting groups.
//
// @OnlyCurrentDoc
//
function onOpen(e) {
  DirectorSheetLib.onOpen(e);
}

function onEdit(e) {
  DirectorSheetLib.onEdit(e);
}

function onToggleDataFilterVisibility() {
  DirectorSheetLib.onToggleDataFilterVisibility();
}

function onGetApplicationData() {
  DirectorSheetLib.onGetApplicationData();
}

function onGenerateSpreadsheetName() {
  DirectorSheetLib.onGenerateSpreadsheetName();
}

function onInsertEmptyRows() {
  DirectorSheetLib.onInsertEmptyRows();
}

function onAboutThis() {
  DirectorSheetLib.onAboutThis();
}

function onApplyUpdates() {
  DirectorSheetLib.onApplyUpdates();
}
