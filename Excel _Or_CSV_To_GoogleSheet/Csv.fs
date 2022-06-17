module Csv

open System
open MyTypes
open GoogleSheet
open FSharp.Data
open ROP_Functions
open System.Threading

//kod je bez validace a (s jednou vyjimkou) bez try-with bloku a slouzi pouze pro vyzkouseni si csv type provider (realny use case je Excel -> Google Sheet)

let csv() =
        
    do printfn "Přesun dat z csv do Google Sheet ..."    

    //nejdrive zrobime typ (viz MyTypes.fs) a pote nacteme nejaky csv soubor (jakykoliv se stejnym typem, nemusi to byt ten samy, jako to nyni mam ja)
    use csvData = MyCSV.Load(@"e:\E\Mirek po osme hodine a o vikendech\Pruvodky Litomerice\CSV Digitalizacni sada\LT-00001 az LT-00204.csv") 
    //use csvData = MyCSV.Load(@"e:\Mirek po osme hodine a o vikendech\LT-00001 az LT-04640.csv")     
    
    //************** Array ponechano vsude zamerne kvuli nutnosti prenosu do DLL C# ************    
    let myHeaderArray =

        let headers = csvData.Headers //csvData.Headers je Array option ...
        match headers with 
        | Some value -> value                                             
        | None       -> do printfn "Varování 3:  Array.Empty" 
                        Array.empty       
        
    let dataRows = 
        let result = csvData.Rows |> Option.ofObj 
        match result with 
               | Some value -> value |> Seq.toArray // ... ale csvData.Rows udelali Seq, zajimave...                                              
               | None       -> do printfn "Varování 4:  Array.Empty" 
                               Array.empty   
     
    //******************** SLOZITY POSTUP **********************
    //pouze pro vyzkouseni si fungovani csv type provider
    let column1  = dataRows |> Array.map(fun row -> row.``Pracovní značení``)
    let column2  = dataRows |> Array.map(fun row -> row.``Digitalizační sada``)
    let column3  = dataRows |> Array.map(fun row -> string row.Archiv)
    let column4  = dataRows |> Array.map(fun row -> string row.Fond)
    let column5  = dataRows |> Array.map(fun row -> string row.``Číslo NAD``)
    let column6  = dataRows |> Array.map(fun row -> string row.``Číslo pomůcky``)
    let column7  = dataRows |> Array.map(fun row -> string row.``Inventární číslo``)
    let column8  = dataRows |> Array.map(fun row -> string row.Signatura)
    let column9  = dataRows |> Array.map(fun row -> string row.``Číslo kartonu``)
    let column10 = dataRows |> Array.map(fun row -> string row.``Upřesňující indentifikátor``)
    let column11 = dataRows |> Array.map(fun row -> string row.Regest)
    let column12 = dataRows |> Array.map(fun row -> string row.``Datace vzniku``)
    let column13 = dataRows |> Array.map(fun row -> string row.Poznámka)   

    let myMultiArray = [|column1; column2; column3; column4; column5; column6; column7; column8; column9; column10; column11; column12; column13|]
    
    //DLL C#
    //do WritingToGoogleSheets.WriteToGoogleSheets(myHeaderArray, myMultiArray, @"c:\Users\Mira\Downloads\nomadic-charge-314614-fd817ff829fa.json", @"1G15Mn_A9EjIXiS0UODz7WVsmx6BoSJGYA4tA-MtoEC4", "Sheet1", 1, 1, true)
 

    //******************** JEDNODUSSI POSTUP ********************
    let myMultiArray =            

        //slo by jeste let mySingleRow = dataRows.GetValue(i) a prohnat to pres Array.mapi nebo Seq pro dany interval
        let myDeconstructedRows = dataRows |> Array.map(fun row -> row.Deconstruct())  
        myDeconstructedRows   
        |> Array.map(fun item ->
                            //a..m => sloupce A..M v Google Sheet
                            let (a,b,c,d,e,f,g,h,i,j,k,l,m) = item                                                        
                            let myArray = 
                                //s vyjimkou sloupcu A a B muze dojit u dalsich csv k type inference na int, proto pretypovani
                                [| a; b; string c; string d; string e; string f; string g; string h; string i; string j; string k; string l; string m |]
                            myArray
                    )

    do printfn "Chviličku strpení (csv) ..."

    let myFunction x = 
        let jsonFileName = @"c:\Users\User\source\repos\FSXAMLjson\nomadic-charge-314614-fd817ff829fb.json" //u tasks nesmi byt stejny json pro excel a csv
        let id = @"1G15Mn_A9EjIXiS0UODz7WVsmx6BoSJGYA4tA-MtoEC4" // je to soucast URL Google tabulky
        let sheetName = "Sheet2"
        do WritingToGoogleSheets.WriteToGoogleSheets(myHeaderArray, myMultiArray, jsonFileName, id, sheetName, 1, 1, false)  //DLL C#  
    let results = 
        //let finalAction x = ()         
        let ropResults = tryWith myFunction (fun x -> ()) (fun ex -> failwith)
        ropResults |> deconstructor () "K přenosu dat z csv do Google Sheet nedošlo"  
    results 
    do Thread.Sleep(0)
    //0 // return an integer exit code


