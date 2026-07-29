# ShellN.Extensions

A friendly, AOT compatible managed layer over some of the Windows Shell namespace APIs, on top of ShellN.

It lets you browse and operate the shell as objects rather than paths, so you reach everything the shell knows about, drives and folders, libraries, archives opened as folders, and virtual locations like This PC.

## Highlights

* navigate with ShellItem, ShellFolder and KnownFolder.
* read a PropertyStore, display names, attributes and icons.
* get a stream for an item, so its content can be read even when it is not a plain file.
* invoke the real Windows context menu and its verbs.
* watch the shell for changes with ChangeNotifier.
