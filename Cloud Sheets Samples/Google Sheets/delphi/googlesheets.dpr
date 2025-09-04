(*
 * Cloud Sheets 2024 Delphi Edition - Sample Project
 *
 * This sample project demonstrates the usage of Cloud Sheets in a 
 * simple, straightforward way. It is not intended to be a complete 
 * application. Error handling and other checks are simplified for clarity.
 *
 * www.nsoftware.com/cloudsheets
 *
 * This code is subject to the terms and conditions specified in the 
 * corresponding product license agreement which outlines the authorized 
 * usage and restrictions.
 *)

program googlesheets;

uses
  Forms,
  googlesheetsf in 'googlesheetsf.pas' {FormGooglesheets};

begin
  Application.Initialize;

  Application.CreateForm(TFormGooglesheets, FormGooglesheets);
  Application.Run;
end.


         
