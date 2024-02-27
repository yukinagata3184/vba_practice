@echo off
@rem decombine command is decombine all of bin/.
@rem So move to past/ that decombined xlsm.
@rem author:yukinagata3184

set "source=bin"
set "destination=past"

robocopy "%source%" "%destination%" /mov

exit /b