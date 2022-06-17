module ROP_Functions

open System
open MyTypes

let tryWith f1 f2 x =
    try
        try
           f1 x |> Success
        finally
           f2 x
    with
    | ex -> Failure ex.Message  

let deconstructor f s =  
    function
    | Success x  -> let y = x
                    y                                
    | Failure ex -> ex |> printfn "Popis chyby: %s"  
                    s  |> printfn "%s"
                    f   
