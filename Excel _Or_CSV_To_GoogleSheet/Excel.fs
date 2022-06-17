module Excel

open System
open MyTypes
open GoogleSheet
open ROP_Functions
open System.Threading
open FSharp.Interop.Excel

open FSharp.Data           //F# SqlCommandProvider, SqlProgrammabilityProvider
open FSharp.Data.Sql       //F# SqlDataProvider
open FSharp.Data.SqlClient //F# ResultType.DataReader 

//************** Array je na vystupu zamerne kvuli nutnosti prenosu do DLL C# ************  
//************** Connection string je zamerne vicekrat ***********************************
let [<Literal>] TypeProviderConn = @"Data Source=Misa\SQLEXPRESS;Initial Catalog=DGSada;Integrated Security=True" 

//DEFINICE CASTO POUZIVANYCH FUNKCI 
let private pressEnterToContinue() = 
    do printfn "Press Enter to continue"                 
    Seq.initInfinite (fun _ -> Console.ReadKey().Key)  
    |> Seq.takeWhile ((<>) ConsoleKey.Enter) //(fun item -> item <> ConsoleKey.Enter)
    |> Seq.iter (fun _ -> ()) 

let private readLine again = 
    //do printfn "Zadej plnou cestu xlsx souboru %s" <| again  
    //Console.ReadLine()
    let str = @"e:\E\Mirek po osme hodine a o vikendech\Pruvodky Litomerice\XLSXnew Digitalizacni sada\LT-00001 az LT-00204 DGSada 01-06-2021.xlsx"
    //let str = @"e:\E\Mirek po osme hodine a o vikendech\Digisady26-04-2021.xlsx"
    str

let private getPath again path firstCycle = 
    match firstCycle with
    | true -> path
    | _    -> readLine <| again 

let private deleteDbData (rows: string[][]) = 
    use cmdDelete = new SqlCommandProvider<"DELETE FROM SOAL", TypeProviderConn>(TypeProviderConn)    
    let recordsAffected = cmdDelete.Execute() 
    do assert ((=) recordsAffected rows.Length)//pouze pro debug

let private condition path = 
    let cond1 = System.IO.File.Exists(path) 
    System.IO.File.Exists(path) && System.IO.Path.GetExtension(path) = ".xlsx" 
         
//ZISKANI DAT Z EXCELU - CAST 1 //VYPADA TO, ZE LEPE SE DOKAZU ZBAVIT NULL U C# EXCELDATAREADER NEZ U F# TYPE PROVIDER PRO EXCEL 
let private getHeadersAndTail path firstCycle =  
           
    let rec doUntil path =                         
        let getExcelFileDataList() = 
            let rec doUntil1 path1 =                 
                let checkExcelFileExist() = 
                    // pri zobecneni to vyhazuje unmanaged exception //TODO zjisti, cemu to tak je                    
                    let result = 
                        let ropResults = 
                            //let finalAction x = ()
                            let myFunction x = condition <| path1          
                            tryWith myFunction (fun x -> ()) (fun ex -> failwith)
                        ropResults |> deconstructor false String.Empty 
                    result
                let checkExcelProvider results = 
                    match results with
                    | true  ->  
                            // kdybych tuto fci zobecnil, vyhodi to unmanaged exception u ExcelProvider //TODO zjisti, cemu to tak je
                            let myFunction x =                                 
                                let result = 
                                    let file = new ExcelProvider(path1) |> Option.ofObj
                                    match file with
                                    | Some value -> let data = value.Data |> Seq.toList
                                                    data 
                                    | None       -> do printfn "Varování 1:  List.Empty"
                                                    List.Empty
                                result                            
                            let results = 
                                let ropResults = tryWith myFunction (fun x -> ()) (fun ex -> failwith)
                                ropResults |> deconstructor List.Empty String.Empty
                            results                              
                    | false -> List.Empty 
                checkExcelFileExist() |> checkExcelProvider
            doUntil1 (getPath <| "ještě jednou (err 6)" <| path <| true)             
        let getHeadAndTailList list = 
            match list with
            | head :: tail ->                     
                            match tail.Length <= 1 with                              
                            | false -> head, tail
                            | true  -> do printfn "Chybný Excel soubor (err 5)"
                                       null, List.Empty //u typu Excel<..> Row neexistuje neco.Empty, bohuzel musime pouzivat null
            | []           -> 
                            do printfn "Chybný Excel soubor (err 4)"  
                            null, List.Empty  
        
        let getHeadAndTail() = getExcelFileDataList() |> getHeadAndTailList 

        let result path =  
            let getHeadAndTail = getHeadAndTail()
            match getHeadAndTail = (null, List.Empty) with //tohle by melo eliminovat null a zajistit vyhnuti se NullReferenceException
            | false -> getHeadAndTail
            | true  -> doUntil (getPath <| "ještě jednou (err 1)" <| path <| false) 
        result path         
    
    doUntil (getPath <| String.Empty <| path <| true) 
           
//ZISKANI DAT Z EXCELU - CAST 2 
let private getExcelData() =   
                      
    let rec doUntil path =  
            match condition path with
            | true  -> getHeadersAndTail <|path <| true
            | false -> doUntil ("ještě jednou (err 2)" |> readLine )                
    doUntil (String.Empty |> readLine)  
 
let readDataFromExcel() =  

    let myFunction x = 
        let (headers, tail) = getExcelData()       
        let dataRows = tail |> List.take (tail.Length - 1) |> List.toArray //quli DLL C# musi byt array, tail.Length - 1 => z neznamych duvodu se dosazuji hodnoty null do posledniho radku Excelu
        let myRange = 
            let numberOfColumns = 13
            [|0..numberOfColumns - 1|]
        let myHeaderArray = 
            myRange |> Array.map(fun i -> (headers.GetValue i) |> string) //s downcast to nefunguje, ale s try-with to funguje s jednoduchym cast       
        let myMultiArray =
            dataRows |> Array.map(fun row -> myRange |> Array.map(fun i -> (row.GetValue i) |> string)) //s downcast to nefunguje, ale s try-with to funguje s jednoduchym cast  
        myHeaderArray, myMultiArray
    let results =
        let ropResults = tryWith myFunction (fun x -> ()) (fun ex -> failwith)
        ropResults |> deconstructor (Array.empty, Array.empty) String.Empty
    results

//HLAVNI FUNKCE - PRENOS DAT Z EXCELU DO GOOGLE SHEETS A DO DB
let excel() = 
    
    do printfn "Přesun dat z xlsx do Google Sheet a db ..."  
    
    // => GOOGLE SHEETS
    let writeDataIntoGoogleSheets readExcel = 
        
        let rec recursiveTryWith() =

            let (header, rows) = readExcel 
            let jsonFileName = @"c:\Users\User\source\repos\FSXAMLjson\nomadic-charge-314614-fd817ff829fa.json" //u tasks nesmi byt stejny json pro excel a csv
            let id = @"1G15Mn_A9EjIXiS0UODz7WVsmx6BoSJGYA4tA-MtoEC4" // je to soucast URL Google tabulky
            let sheetName = "Sheet1"

            try
                try
                   do printfn "Chviličku strpení (excel) ..."
                   do WritingToGoogleSheets.WriteToGoogleSheets(
                                                                 header,
                                                                 rows, 
                                                                 jsonFileName, 
                                                                 id, 
                                                                 sheetName,
                                                                 1, 1, false
                                                               ) //DLL C# 
                   true
                finally
                   ()                                
            with
            | ex ->          
                    ex.Message |> printfn "Popis chyby: %s"
                    do printfn "No jo, chyba. Tož ji odstraň."
                    do pressEnterToContinue()
                    false                            
       
        Seq.initInfinite (fun _ -> recursiveTryWith())
        |> Seq.skipWhile ((=) false)
        |> Seq.head |> ignore //while block pro non-unit results, ja ale v danem pripade potrebuji na vystupu unit, proto |> ignore
      
        (*
            //to same lze takto (ale jen pro unit)://
            Seq.initInfinite (fun _ -> recursiveTryWith()) 
            |> Seq.takeWhile ((=) false) 
            |> Seq.iter (fun _ -> ())  
        *)

    // => DATABASE
    let writeDataIntoDb (readExcel: string[]*string[][]) = //neumi dedukovat typ 
            
            let (header, rows) = readExcel             

            // 1) vyzkousime SqlCommandProvider
            let doItWithSqlCommandProvider() = 

                do deleteDbData <| rows
               
                let myFunction x = 

                    use cmdInsert = new SqlCommandProvider<"
                        INSERT INTO SOAL (Id,
                                         [Pracovní značení],
                                         [Digitalizační sada],
                                         [Archiv],
                                         [Fond],
                                         [Číslo NAD],
                                         [Číslo pomůcky],
                                         [Inventární číslo],
                                         [Signatura],
                                         [Číslo kartonu],
                                         [Upřesňující indentifikátor],
                                         [Regest],
                                         [Datace vzniku],
                                         [Poznámka])
                        VALUES (@val01, @val02, @val03, @val04, @val05, @val06, @val07,
                                @val08, @val09, @val10, @val11, @val12, @val13, @val14)     
                        ", TypeProviderConn>(TypeProviderConn)                    
                
                    rows 
                    |> Array.iteri(fun i row -> 
                                             cmdInsert.Execute(val01 = i + 1, val02 = (row |> Array.item 0), val03 = row.[1],
                                                               val04 = row.[2], val05 = row.[3], val06 = row.[4], 
                                                               val07 = row.[5], val08 = row.[6], val09 = row.[7], 
                                                               val10 = row.[8], val11 = row.[9], val12 = row.[10], 
                                                               val13 = row.[11], val14 = row.[12]) |> ignore 
                                   )  
                                   
                    
                    //Add je proste insert, prida polozku hned za tu posledni
                    let str = @"N/A"
                    use cmdInsert = new SqlCommandProvider<"
                        INSERT INTO SOAL (Id,
                                            [Pracovní značení],
                                            [Digitalizační sada],
                                            [Archiv],
                                            [Fond],
                                            [Číslo NAD],
                                            [Číslo pomůcky],
                                            [Inventární číslo],
                                            [Signatura],
                                            [Číslo kartonu],
                                            [Upřesňující indentifikátor],
                                            [Regest],
                                            [Datace vzniku],
                                            [Poznámka])
                        VALUES (@val01, @val02, @val03, @val04, @val05, @val06, @val07,
                                @val08, @val09, @val10, @val11, @val12, @val13, @val14)     
                        ", TypeProviderConn>(TypeProviderConn)                    
                                
                    cmdInsert.Execute(val01 = 205, val02 = str, val03 = str,
                                        val04 = str, val05 = str, val06 = str, 
                                        val07 = str, val08 = str, val09 = str, 
                                        val10 = str, val11 = str, val12 = str, 
                                        val13 = str, val14 = str) |> ignore                         
                       
                                   
                let results = 
                    let ropResults = tryWith myFunction (fun x -> ()) (fun ex -> failwith)
                    ropResults |> deconstructor () String.Empty
                results               
            
            (*
            // 2) vyzkousime SqlProgrammabilityProvider
            let doItWithSqlProgrammabilityProvider() = 
                
                let myFunction x = 
                    
                    do deleteDbData <| rows

                    use myTableData = new MySOALTable.dbo.Tables.SOAL() //pri pouziti SqlProgrammabilityProvider
                    use loadTableData = new SqlCommandProvider<"SELECT * FROM SOAL", TypeProviderConn, ResultType.DataReader>(TypeProviderConn)
                    do loadTableData.Execute() |> myTableData.Load
                
                    rows
                    |> Array.iteri(fun i row -> 
                                             let newRow = 
                                                 myTableData.NewRow(i + 1, Some row.[0], Some row.[1], 
                                                                    Some row.[2], Some row.[3], Some row.[4], 
                                                                    Some row.[5], Some row.[6], Some row.[7], 
                                                                    Some row.[8], Some row.[9], Some row.[10], 
                                                                    Some row.[11], Some row.[12])               
                                             myTableData.Rows.Add <| newRow
                                  ) |> ignore
                    let recordsAffected = myTableData.Update() 
                    assert(recordsAffected = rows.Length)                
                
                let results = 
                    let ropResults = tryWith myFunction (fun x -> ()) (fun ex -> failwith)
                    ropResults |> deconstructor () String.Empty
                results                  
             *)
            // 3) vyzkousime SqlDataProvider //ORM //nejake pomale je to...
            let doItWithSqlDataProvider() = 
                    
                let myFunction x = 
                    
                    let context = ContextClass.GetDataContext()
                    
                    //nefunguje
                    //let deleteRow = context.Dbo.Soal.Create()
                    //deleteRow.Delete()
                    //context.SubmitUpdates()                  
                    
                    do deleteDbData <| rows

                    rows
                    |> Array.iteri(fun i row ->                                              
                                                let newRow = context.Dbo.Soal.Create() //tady je jmeno tabulky
                                                newRow.Id                        <- i + 1
                                                newRow.PracovníZnačení           <- Some row.[0]
                                                newRow.DigitalizačníSada         <- Some row.[1]
                                                newRow.Archiv                    <- Some row.[2]
                                                newRow.Fond                      <- Some row.[3]
                                                newRow.ČísloNad                  <- Some row.[4]
                                                newRow.ČísloPomůcky              <- Some row.[5]
                                                newRow.InventárníČíslo           <- Some row.[6]
                                                newRow.Signatura                 <- Some row.[7]
                                                newRow.ČísloKartonu              <- Some row.[8]
                                                newRow.UpřesňujícíIndentifikátor <- Some row.[9]
                                                newRow.Regest                    <- Some row.[10]
                                                newRow.DataceVzniku              <- Some row.[11]
                                                newRow.Poznámka                  <- Some row.[12] 
                                                context.SubmitUpdates()
                                    ) |> ignore                    
                
                let results = 
                    let ropResults = tryWith myFunction (fun x -> ()) (fun ex -> failwith)
                    ropResults |> deconstructor () String.Empty
                results 
            
            doItWithSqlCommandProvider()  
            //doItWithSqlProgrammabilityProvider()
            //doItWithSqlDataProvider()
   
    let readDataFromExcel = readDataFromExcel()
    readDataFromExcel |> writeDataIntoGoogleSheets 
    readDataFromExcel |> writeDataIntoDb 
    
    //do Thread.Sleep(5000)
    //0 // return an integer exit code
