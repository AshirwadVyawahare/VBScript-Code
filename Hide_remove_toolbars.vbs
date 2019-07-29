set x = createobject("internetexplorer.application")

'Removes tool bar, address bar, resize and close buttons
x.navigate2 "www.google.com" : x.fullscreen = true : x.visible = True

'Remove toolbar, status bar, addressbar
x.navigate2 "www.google.com" : x.toolbar = false _
: x.menubar = false : x.statusbar = false : x.visible = True


set x = nothing