using XLSX, DataFrames, Statistics, Dates, MixedModels, CairoMakie, Random, StatsBase

df  = DataFrame(XLSX.readtable("data/bleph-cw-v2.xlsx", 2))
rename!(df,
    "Study #" => "individual",
    "Limbus to Limbus measured with ruler on Servoy (mm) T0" => "LLT0",
    "OD MRD1 measured with ruler T0" => "ODMRD1T0",
    "OD MRD1 converted T0" => "ODMRD1cT0",
    "OD TPS measured with ruler T0" => "ODTPST0",
    "OD TPS converted T0" => "ODTPScT0",
    "OD BFS measured T0" => "ODBFST0",
    "OD BFS converted T0" => "ODBFScT0",
    "OS MRD1 measured with ruler T0" => "OSMRD1T0",
    "OS MRD1 converted T0" => "OSMRD1cT0",
    "OS TPS measured with ruler T0" => "OSTPST0",
    "OS TPS converted T0" => "OSTPScT0",
    "OS BFS measured T0" => "OSBFST0",
    "OS BFS converted T0" => "OSBFScT0",
    "Limbus to Limbus measured with ruler on Servoy (mm) T1" => "LLT1",
    "OD MRD1 measured with ruler PostOp T1" => "ODMRD1T1",
    "OD MRD1 converted PostOp T1" => "ODMRD1cT1",
    "OD TPS measured with ruler PostOp T1" => "ODTPST1",
    "OD TPS converted PostOp T1" => "ODTPScT1",
    "OD BFS measured PostOp T1" => "ODBFST1", 
    "OD BFS converted PostOp T1" => "ODBFScT1",
    "OS MRD1 measured with ruler PostOp T1" => "OSMRD1T1",
    "OS MRD1 converted PostOp T1" => "OSMRD1cT1",
    "OS TPS measured with ruler PostOp T1" => "OSTPST1",
    "OS TPS converted PostOp T1" => "OSTPScT1",
    "OS BFS measured PostOp T1" => "OSBFST1", 
    "OS BFS converted PostOp T1" => "OSBFScT1",
    "Limbus to Limbus measured with ruler on Servoy (mm) T2" => "LLT2",
    "OD MRD1 measured with ruler POY1 T2" => "ODMRD1T2",
    "OD MRD1 converted POY1 T2" => "ODMRD1cT2",
    "OD TPS measured with ruler POY1 T2" => "ODTPST2",
    "OD TPS converted POY1 T2" => "ODTPScT2",
    "OD BFS measured POY1 T2" => "ODBFST2",
    "OD BFS converted POY1 T2" => "ODBFScT2",
    "OS MRD1 measured with ruler POY1 T2" => "OSMRD1T2",
    "OS MRD1 converted POY1 T2" => "OSMRD1cT2",
    "OS TPS measured with ruler POY1 T2" => "OSTPST2",
    "OS TPS converted POY1 T2" => "OSTPScT2",
    "OS BFS measured POY1 T2" => "OSBFST2",
    "OS BFS converted POY1 T2" => "OSBFScT2",
    "Limbus to Limbus measured with ruler on Servoy (mm) T3" => "LLT3",
    "OD MRD1 measured with ruler T3" => "ODMRD1T3",
    "OD MRD1 converted T3" => "ODMRD1cT3",
    "OD TPS measured with ruler T3" => "ODTPST3",
    "OD TPS converted T3" => "ODTPScT3",
    "OD BFS measured T3" => "ODBFST3",
    "OD BFS converted T3" => "ODBFScT3",
    "OS MRD1 measured with ruler T3" => "OSMRD1T3",
    "OS MRD1 converted T3" => "OSMRD1cT3",
    "OS TPS measured with ruler T3" => "OSTPST3",
    "OS TPS converted T3" => "OSTPScT3",
    "OS BFS measured T3" => "OSBFST3",
    "OS BFS converted T3" => "OSBFScT3",
    "Limbus to Limbus measured with ruler on Servoy (mm) T4" => "LLT4",
    "OD MRD1 measured with ruler T4" => "ODMRD1T4",
    "OD MRD1 converted T4" => "ODMRD1cT4",
    "OD TPS measured with ruler T4" => "ODTPST4",
    "OD TPS converted T4" => "ODTPScT4",
    "OD BFS measured T4" => "ODBFST4",
    "OD BFS converted T4" => "ODBFScT4",
    "OS MRD1 measured with ruler T4" => "OSMRD1T4",
    "OS MRD1 converted T4" => "OSMRD1cT4",
    "OS TPS measured with ruler T4" => "OSTPST4",
    "OS TPS converted T4" => "OSTPScT4",
    "OS BFS measured T4" => "OSBFST4",
    "OS BFS converted T4" => "OSBFScT4",
    "Limbus to Limbus measured with ruler on Servoy (mm) T5" => "LLT5",
    "OD MRD1 measured with ruler T5" => "ODMRD1T5",
    "OD MRD1 converted T5" => "ODMRD1cT5",
    "OD TPS measured with ruler T5" => "ODTPST5",
    "OD TPS converted T5" => "ODTPScT5",
    "OD BFS measured T5" => "ODBFST5",
    "OD BFS converted T5" => "ODBFScT5",
    "OS MRD1 measured with ruler T5" => "OSMRD1T5",
    "OS MRD1 converted T5" => "OSMRD1cT5",
    "OS TPS measured with ruler T5" => "OSTPST5",
    "OS TPS converted T5" => "OSTPScT5",
    "OS BFS measured T5" => "OSBFST5",
    "OS BFS converted T5" => "OSBFScT5",
    "Limbus to Limbus measured with ruler on Servoy (mm) T6" => "LLT6",
    "OD MRD1 measured with ruler T6" => "ODMRD1T6",
    "OD MRD1 converted T6" => "ODMRD1cT6",
    "OD TPS measured with ruler T6" => "ODTPST6",
    "OD TPS converted T6" => "ODTPScT6",
    "OD BFS measured T6" => "ODBFST6",
    "OD BFS converted T6" => "ODBFScT6",
    "OS MRD1 measured with ruler T6" => "OSMRD1T6",
    "OS MRD1 converted T6" => "OSMRD1cT6",
    "OS TPS measured with ruler T6" => "OSTPST6",
    "OS TPS converted T6" => "OSTPScT6",
    "OS BFS measured T6" => "OSBFST6",
    "OS BFS converted T6" => "OSBFScT6",
    )

rename!(df, ["T1", "T2", "T3", "T4", "T5", "T6"] .* " Time" .=> "Δ" .* ["T1", "T2", "T3", "T4", "T5", "T6"])
findall(==(0), [typeof(df[i, "T0"]) == Date for i in 1:nrow(df)])
findall(==(0), [typeof(df[i, "T1"]) == Date for i in 1:nrow(df)])
df[44, "T1"] = Date(2001, 8, 16)
findall(==(0), [typeof(df[i, "T2"]) == Date || ismissing(df[i, "T2"]) for i in 1:nrow(df)])
findall(==(0), [typeof(df[i, "T3"]) == Date || ismissing(df[i, "T3"]) for i in 1:nrow(df)])
findall(==(0), [typeof(df[i, "T4"]) == Date || ismissing(df[i, "T4"]) for i in 1:nrow(df)])
findall(==(0), [typeof(df[i, "T5"]) == Date || ismissing(df[i, "T5"]) for i in 1:nrow(df)])
findall(==(0), [typeof(df[i, "T6"]) == Date || ismissing(df[i, "T6"]) for i in 1:nrow(df)])

s = 0
for i in 1:nrow(df)
    for col in ["T2", "T3", "T4", "T5", "T6"]
        if !ismissing(df[i, col])
            d1 = Dates.days(df[i, col] - df[i, "T0"]) / 365
            d2 = df[i, "Δ" * col]
            if d1 == d2
                s += 1
                @info "Discrepancy at $col for $(i)th row with $d1, $d2"
            end
        end
    end
end

for i in 1:nrow(df)
    for col in ["T1", "T2", "T3", "T4", "T5", "T6"]
        if ismissing(df[i, col])
            df[i, "Δ" * col] = missing
        else
            df[i, "Δ" * col] = Dates.days(df[i, col] - df[i, "T0"]) / 365
        end
    end
end

pairwise(cor, eachcol([df[!, "ODMRD1T0"] df[!, "ODTPST0"] df[!, "ODBFST0"]]), skipmissing = :pairwise)
pairwise(cor, eachcol([df[!, "ODMRD1cT0"] df[!, "ODTPScT0"] df[!, "ODBFScT0"]]), skipmissing = :pairwise)
