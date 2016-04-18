:: Converts *.doc files to *.pdf
::
:: Runs doc2pdf.vbs located in J:\sc performing the actual conversion
:: Moves the finished pdfs to another folder
:: Deletes the doc files.
:: Requirements: needs doc2pdf.vbs in J:\sc
::
:: _File structure_
::
:: J:\rec\suggestions
::						\pdf
::
:: If this is on a UNC path, use the PUSHD and POPD commands.
:: Can be run from cmd line or called from another script.
:: Does not take any arguments

@ECHO OFF
ECHO (C) J. Ditzel & ECHO.May 2015
SETLOCAL

	SET path=%PATH%;J:\sc
	SET FPATH="J:\rec\suggestions"
	SET CURD=%CD%
	
	PUSHD J:\
	CD %FPATH%
		FOR /R %FPATH% %%G IN (*.doc) DO (
			doc2pdf.vbs "%%G" 
			FOR /R %FPATH% %%H IN (*.pdf) DO MOVE /Y "%%H" %FPATH%\pdf >nul	)
		DEL /Q *.doc
:: Clean up
	POPD
	CD %CURD%	
ENDLOCAL
ECHO Completed successfully.
@ECHO ON