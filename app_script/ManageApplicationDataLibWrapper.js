//
// Copyright 2023 Douglas Herrick
//
// Use of this source code is governed by an MIT-style
// license that can be found in the LICENSE file or at
//
// https://opensource.org/licenses/MIT.
//
// This collection of wrapper functions belay the interesting work to the library that manages
// the application data.
//
// @OnlyCurrentDoc
//
function onOpen(e) {
  ManageApplicationDataLib.onOpen(e);
}

function onEdit(e) {
  ManageApplicationDataLib.onEdit(e);
}

function onToggleDataFilterVisibility() {
  ManageApplicationDataLib.onToggleDataFilterVisibility();
}

function onGetApplicationData(sheet, rows) {
  ManageApplicationDataLib.onGetApplicationData(sheet, rows);
}

function onGenerateSpreadsheetName(sheet) {
  ManageApplicationDataLib.onGenerateSpreadsheetName(sheet);
}

function onAboutThis() {
  ManageApplicationDataLib.onAboutThis();
}
