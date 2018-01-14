# dictationbridge-jfw

This is the Jaws scripts package for DictationBridge. It depends on the [screen-reader-independent core](https://github.com/dictationbridge/dictationbridge-core).

Support for Dragon NaturallySpeaking is still under heavy development.

## Building

To build this package, you need Python 2.7, a recent version of SCons, [NSIS](http://nsis.sourceforge.net/Main_Page), and Visual Studio 2015 or 2017. If using Visual Studio 2017, be sure to install MFC/ATL tools.

First, fetch Git submodules with this command:

    git submodule update --init --recursive

Then run:

```
scons
makensis installer.nsi
```

This will produce an installer as `DictationBridge.exe`.
