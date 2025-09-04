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

program office365excel;

uses
  Forms,
  office365excelf in 'office365excelf.pas' {FormOffice365excel};

begin
  Application.Initialize;

  Application.CreateForm(TFormOffice365excel, FormOffice365excel);
  Application.Run;
end.


         
