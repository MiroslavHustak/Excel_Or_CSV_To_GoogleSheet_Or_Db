module StartForTestingPurposes

open System

open Csv
open Excel
open MyTypes
open ROP_Functions
open System.Threading.Tasks

// ASYNCHRONOUS @ MULTITHREADING
// REDUNDANT CODE (for testing and learning purposes only)
// Pouze s fixne stanovenou cestou k xlsx souboru

let private str = "HH:mm:ss"

let private processStart() =     
    let processStartTime = $"Začátek procesu: {DateTime.Now.ToString(str)}"
    Console.WriteLine(processStartTime)

let private processEnd() =     
    let processEndTime = $"Konec procesu: {DateTime.Now.ToString(str)}"
    Console.WriteLine(processEndTime)

let start1() = 

    //Single thread  //doba trvani t + 6 az 9
    do processStart()
    do csv() 
    do excel() 
    do processEnd()  

    //Tasks //doba trvani t (14 vterin pro 4640 radku) - nejrychlejsi ze vseho
    //Pouziva se pole, proto stejny typ kazde polozky!!!
    do processStart()
    let ts = [| 
                Task.Factory.StartNew(fun () -> csv())   //pozor na jsonFileName, nesmi se otevirat ten samy soubor
                Task.Factory.StartNew(fun () -> excel()) //pozor na jsonFileName, nesmi se otevirat ten samy soubor
             |]
    Task.WaitAll(ts |> Seq.cast<Task> |> Array.ofSeq)
    do processEnd()
  
    //let myTaskFunctionDU x = 
    let myTaskFunction() = 

        let task1 param = 
            async
                {
                  //do! Async.Sleep 0
                  return CsvInt param
                }        

        let task2 param = 
            async 
                { 
                  //do! Async.Sleep 0
                  return ExcelInt param 
                }
 
        let du: TaskResults[] = [| task1 (csv()); task2 (excel()) |] 
                                |> Async.Parallel 
                                |> Async.Catch
                                |> Async.RunSynchronously
                                |> function
                                   | Choice1Of2 result    -> result
                                   | Choice2Of2 (ex: exn) -> do printfn "Popis chyby async: %s" <| ex.Message
                                                             Array.empty  
        du |> ignore 
    
    //Asynchronous workflows (RunSynchronously, pokud potrebujeme value na vystupu) //doba trvani t + 8 az 15 vterin
    do processStart()
    myTaskFunction()
    //let ropResults() = tryWith myTaskFunctionDU (fun x -> ()) (fun ex -> failwith)             
    //ropResults() |> deconstructor |> ignore
    do processEnd()

    //let myTaskFunction x =
    let myTaskFunction2() =
        
        async
            {
              //do! 0 |> Async.Sleep 
              return csv()
            } |> Async.Start //pouze pro unit na vystup   
           
        async 
            {
              //do! 0 |> Async.Sleep 
              return excel() 
            } |> Async.Start  //pouze pro vysledek unit    
    
    do processStart() //tady nerelevantni, bo asynchronous... :-)
    myTaskFunction2()
    //let ropResults() = tryWith myTaskFunction (fun x -> ()) (fun ex -> failwith)             
    //ropResults() |> deconstructor |> ignore
    do processEnd() //tady nerelevantni, bo asynchronous... :-)                    
    
    0 // return an integer exit code
