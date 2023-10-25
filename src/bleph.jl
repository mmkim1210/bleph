using XLSX, DataFrames, Statistics, Dates, MixedModels, CairoMakie, Random, StatsBase, Printf, HypothesisTests, Distributions, GLM

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

df[19, "T3"] = missing
df[19, "ΔT3"] = missing
df[53, "T4"] = missing
df[53, "ΔT4"] = missing
df[56, "T5"] = missing
df[56, "ΔT5"] = missing

pairwise(cor, eachcol([df[!, "ODMRD1T0"] df[!, "ODTPST0"] df[!, "ODBFST0"]]), skipmissing = :pairwise)
pairwise(cor, eachcol([df[!, "ODMRD1cT0"] df[!, "ODTPScT0"] df[!, "ODBFScT0"]]), skipmissing = :pairwise)

for i in ["MRD1", "TPS", "BFS"]
    for j in ["T0", "T1"]
        for k in ["OD", "OS"]
            df[!, k * i * "c" * j] = convert(Vector{Float64}, df[!, k * i * "c" * j])
            df[!, k * i * j] = convert(Vector{Float64}, df[!, k * i * j])
        end
    end
end

idx = findfirst(>(50), df[!, "ODTPScT1"])
df[idx, "ODTPScT1"] = df[idx, "OSTPScT1"]

newdf = DataFrame(individual = Float64[], ΔT = Float64[], ODMRD1c = Float64[], ODTPSc = Float64[], ODBFSc = Float64[])
for i in 1:nrow(df)
    for j in 1:6
        if !ismissing(df[i, "ΔT$j"])
            push!(newdf, [df[i, "individual"] df[i, "ΔT$j"] df[i, "ODMRD1cT$j"] df[i, "ODTPScT$j"] df[i, "ODBFScT$j"]])
        end
    end
end

begin
    f = Figure()
    axs = [Axis(f[i, j]) for i in 2:3, j in 1:3]
    for (j, measure) in enumerate(["MRD1", "TPS", "BFS"])
        for (i, time) in enumerate(["T0", "T1"])
            scatter!(axs[i, j], df[!, "OD" * measure * "c" * time], df[!, "OS" * measure * "c" * time], color = ("black", 0.5))
            text!(axs[i, j], maximum(df[!, "OD" * measure * "c" * time]), minimum(df[!, "OS" * measure * "c" * time]), 
                text = "R = $(round(cor(df[!, "OD" * measure * "c" * time], df[!, "OS" * measure * "c" * time]); sigdigits = 2))", 
                align = (:right, :bottom))
        end
    end
    [hidedecorations!(axs[i, j], ticklabels = false, ticks = false) for i in 1:2, j in 1:3]
    titles = ["MRD1", "TPS1", "BFS"]
    for (i, title) in enumerate(titles)
        Box(f[1, i], color = :gray90)
        Label(f[1, i], title, tellwidth = false, padding = (3, 3, 3, 3))
    end
    Box(f[2, 4], color = :gray90)
    Label(f[2, 4], "Pre-op", tellheight = false, rotation = -pi / 2, padding = (3, 3, 3, 3))
    Box(f[3, 4], color = :gray90)
    Label(f[3, 4], "Post-op", tellheight = false, rotation = -pi / 2, padding = (3, 3, 3, 3))
    rowgap!(f.layout, 1, 0)
    colgap!(f.layout, 3, 0)
    Label(f[2:3, 0], text = "OS measurements", rotation = pi / 2)
    Label(f[4, 1:3], text = "OD measurements")
    save("figs/OD-OS-correlation.pdf", f)
    f
end

begin
    f = Figure()
    axs = [Axis(f[i, j]) for i in 1:2, j in 1:3]
    for (i, time) in enumerate(["T0", "T1"])
        scatter!(axs[i, 1], df[!, "ODMRD1c" * time], df[!, "ODTPSc" * time], color = ("black", 0.5))
        scatter!(axs[i, 2], df[!, "ODMRD1c" * time], df[!, "ODBFSc" * time], color = ("black", 0.5))
        scatter!(axs[i, 3], df[!, "ODTPSc" * time], df[!, "ODBFSc" * time], color = ("black", 0.5))
        text!(axs[i, 1], maximum(df[!, "ODMRD1c" * time]), maximum(df[!, "ODTPSc" * time]), 
            text = "R = $(round(cor(df[!, "ODMRD1c" * time], df[!, "ODTPSc" * time]); sigdigits = 2))", 
            align = (:right, :top))
        text!(axs[i, 2], maximum(df[!, "ODMRD1c" * time]), maximum(df[!, "ODBFSc" * time]), 
            text = "R = $(round(cor(df[!, "ODMRD1c" * time], df[!, "ODBFSc" * time]); sigdigits = 2))", 
            align = (:right, :top))
        text!(axs[i, 3], maximum(df[!, "ODTPSc" * time]), maximum(df[!, "ODBFSc" * time]), 
            text = "R = $(round(cor(df[!, "ODTPSc" * time], df[!, "ODBFSc" * time]); sigdigits = 2))", 
            align = (:right, :top))
    end
    [hidedecorations!(axs[i, j], ticklabels = false, ticks = false) for i in 1:2, j in 1:3]
    Label(f[1:2, 1, Left()], "OD TPS", rotation = pi / 2, padding = (0, 40, 0, 0))
    Label(f[2, 1, Bottom()], "OD MRD1", padding = (0, 0, 0, 40))
    Label(f[1:2, 2, Left()], "OD BFS", rotation = pi / 2, padding = (0, 40, 0, 0))
    Label(f[2, 2, Bottom()], "OD MRD1", padding = (0, 0, 0, 40))
    Label(f[1:2, 3, Left()], "OD BFS", rotation = pi / 2, padding = (0, 40, 0, 0))
    Label(f[2, 3, Bottom()], "OD TPS", padding = (0, 0, 0, 40))
    Box(f[1, 4], color = :gray90)
    Label(f[1, 4], "Pre-op", tellheight = false, rotation = -pi / 2, padding = (3, 3, 3, 3))
    Box(f[2, 4], color = :gray90)
    Label(f[2, 4], "Post-op", tellheight = false, rotation = -pi / 2, padding = (3, 3, 3, 3))
    colgap!(f.layout, 3, 0)
    save("figs/MRD1-TPS-BFS-correlation.pdf", f)
    f
end

begin
    f = Figure()
    axs = [Axis(f[i, j]) for i in 2:2, j in 1:3]
    for (j, measure) in enumerate(["MRD1", "TPS", "BFS"])
        scatter!(axs[1, j], repeat([1], nrow(df)), df[!, "OD" * measure * "c" * "T0"], color = ("black", 0.5))
        scatter!(axs[1, j], repeat([2], nrow(df)), df[!, "OD" * measure * "c" * "T1"], color = ("black", 0.5))
        for k in 1:nrow(df)
            lines!(axs[1, j], [1, 2], [df[k, "OD" * measure * "c" * "T0"], df[k, "OD" * measure * "c" * "T1"]], color = ("red", 0.3))
        end
        p = @sprintf "%.2E" pvalue(SignedRankTest(df[!, "OD" * measure * "c" * "T0"], df[!, "OD" * measure * "c" * "T1"]))
        text!(axs[1, j], 2.15, 0, text = "p = " * p, align = (:right, :top))
        Box(f[1, j], color = :gray90)
        Label(f[1, j], measure, tellwidth = false, padding = (3, 3, 3, 3))
    end
    [xlims!(axs[i, j], 0.75, 2.25) for i in 1:1, j in 1:3]
    [axs[i, j].xticks = ([1, 2], ["Pre-op", "Post-op"]) for i in 1:1, j in 1:3]
    [hidexdecorations!(axs[i, j], ticklabels = false) for i in 1:1, j in 1:3]
    rowgap!(f.layout, 1, 0)
    save("figs/MRD1-TPS-BFS-post-op-short-term-changes.pdf", f)
    f
end

patient = unique(newdf.individual)
patient_subset = Float64[]
for p in patient
    if count(==(p), newdf.individual) > 2
        push!(patient_subset, p)
    end
end

begin
    m1 = lm(@formula(ODMRD1c ~ ΔT), newdf)
    m2 = lm(@formula(ODTPSc ~ ΔT), newdf)
    m3 = lm(@formula(ODBFSc ~ ΔT), newdf)
    mall = [coef(m1)[1] coef(m1)[2];
        coef(m2)[1] coef(m2)[2];
        coef(m3)[1] coef(m3)[2]]
end

begin
    f = Figure()
    axs = [Axis(f[i, j]) for i in 2:2, j in 1:3]
    for (j, measure) in enumerate(["MRD1", "TPS", "BFS"])
        scatter!(axs[1, j], newdf[!, "ΔT"], newdf[!, "OD" * measure * "c"], color = ("black", 0.5))
        Box(f[1, j], color = :gray90)
        Label(f[1, j], measure, tellwidth = false, padding = (3, 3, 3, 3))
        ablines!(axs[1, j], mall[j, 1], mall[j, 2], color = :gold)
    end
    rowgap!(f.layout, 1, 0)
    save("figs/MRD1-TPS-BFS-post-op-long-term-changes.pdf", f)
    f
end

m1 = fit(MixedModel, @formula(ODBFSc ~ 1 + ΔT + (1 + ΔT|individual)), newdf)
m0 = fit(MixedModel, @formula(ODBFSc ~ 1 + (1 + ΔT|individual)), newdf)

ccdf(Chisq(dof(m1) - dof(m0)), 2 * (loglikelihood(m1) - loglikelihood(m0))) # LRT

m1 = fit(MixedModel, @formula(ODMRD1c ~ 1 + ΔT + (1 + ΔT|individual)), newdf)
m0 = fit(MixedModel, @formula(ODMRD1c ~ 1 + (1 + ΔT|individual)), newdf)
raneftables(m1)[1]

ccdf(Chisq(dof(m1) - dof(m0)), 2 * (loglikelihood(m1) - loglikelihood(m0))) # LRT

m1 = fit(MixedModel, @formula(ODTPSc ~ 1 + ΔT + (1 + ΔT|individual)), newdf)
m0 = fit(MixedModel, @formula(ODTPSc ~ 1 + (1 + ΔT|individual)), newdf)

ccdf(Chisq(dof(m1) - dof(m0)), 2 * (loglikelihood(m1) - loglikelihood(m0))) # LRT

begin
    f = Figure()
    axs = [Axis(f[i, j]) for i in 1:5, j in 1:5]
    for i in 1:5
        for j in 1:5
            ind = 5 * (i - 1) + j
            storage = filter(row -> row.individual == patient_subset[ind], newdf)
            scatter!(axs[i, j], storage.ΔT, storage.ODMRD1c, color = "#4063D8")
            ablines!(axs[i, j], mall[1, 1], mall[1, 2], color = :gold)
            xlims!(axs[i, j], -1, 21)
            ylims!(axs[i, j], minimum(newdf.ODMRD1c) - 1, maximum(newdf.ODMRD1c) + 1)
            m = lm(@formula(ODMRD1c ~ ΔT), storage)
            ablines!(axs[i, j], coef(m)[1], coef(m)[2], color = "#389826")
            m1 = fit(MixedModel, @formula(ODMRD1c ~ 1 + ΔT + (1 + ΔT|individual)), newdf)
            idx = findfirst(==(patient_subset[ind]), df.individual)
            ablines!(axs[i, j], ranef(m1)[1][1, idx] + coef(m1)[1], ranef(m1)[1][2, idx] + coef(m1)[2], color = "#CB3C33")
        end
    end
    [hidedecorations!(axs[i, j], ticklabels = false, ticks = false) for i in 1:5, j in 1:5]
    Label(f[1:5, 0], text = "MRD1", rotation = pi / 2)
    Label(f[6, 1:5], text = "Time (years)")
    resize_to_layout!(f)
    save("figs/MRD1-post-op-long-term-changes-lmm.pdf", f)
    f
end

begin
    f = Figure()
    axs = [Axis(f[i, j]) for i in 1:5, j in 1:5]
    for i in 1:5
        for j in 1:5
            ind = 5 * (i - 1) + j
            storage = filter(row -> row.individual == patient_subset[ind], newdf)
            scatter!(axs[i, j], storage.ΔT, storage.ODTPSc, color = "#4063D8")
            ablines!(axs[i, j], mall[2, 1], mall[2, 2], color = :gold)
            xlims!(axs[i, j], -1, 21)
            ylims!(axs[i, j], minimum(newdf.ODTPSc) - 1, maximum(newdf.ODTPSc) + 1)
            m = lm(@formula(ODTPSc ~ ΔT), storage)
            ablines!(axs[i, j], coef(m)[1], coef(m)[2], color = "#389826")
            m1 = fit(MixedModel, @formula(ODTPSc ~ 1 + ΔT + (1 + ΔT|individual)), newdf)
            idx = findfirst(==(patient_subset[ind]), df.individual)
            ablines!(axs[i, j], ranef(m1)[1][1, idx] + coef(m1)[1], ranef(m1)[1][2, idx] + coef(m1)[2], color = "#CB3C33")
        end
    end
    [hidedecorations!(axs[i, j], ticklabels = false, ticks = false) for i in 1:5, j in 1:5]
    Label(f[1:5, 0], text = "TPS", rotation = pi / 2)
    Label(f[6, 1:5], text = "Time (years)")
    resize_to_layout!(f)
    save("figs/TPS-post-op-long-term-changes-lmm.pdf", f)
    f
end

begin
    f = Figure()
    axs = [Axis(f[i, j]) for i in 1:5, j in 1:5]
    for i in 1:5
        for j in 1:5
            ind = 5 * (i - 1) + j
            storage = filter(row -> row.individual == patient_subset[ind], newdf)
            scatter!(axs[i, j], storage.ΔT, storage.ODBFSc, color = "#4063D8")
            ablines!(axs[i, j], mall[3, 1], mall[3, 2], color = :gold)
            xlims!(axs[i, j], -1, 21)
            ylims!(axs[i, j], minimum(newdf.ODBFSc) - 1, maximum(newdf.ODBFSc) + 1)
            m = lm(@formula(ODBFSc ~ ΔT), storage)
            ablines!(axs[i, j], coef(m)[1], coef(m)[2], color = "#389826")
            m1 = fit(MixedModel, @formula(ODBFSc ~ 1 + ΔT + (1 + ΔT|individual)), newdf)
            idx = findfirst(==(patient_subset[ind]), df.individual)
            ablines!(axs[i, j], ranef(m1)[1][1, idx] + coef(m1)[1], ranef(m1)[1][2, idx] + coef(m1)[2], color = "#CB3C33")
        end
    end
    [hidedecorations!(axs[i, j], ticklabels = false, ticks = false) for i in 1:5, j in 1:5]
    Label(f[1:5, 0], text = "BFS", rotation = pi / 2)
    Label(f[6, 1:5], text = "Time (years)")
    resize_to_layout!(f)
    save("figs/BFS-post-op-long-term-changes-lmm.pdf", f)
    save("figs/BFS-post-op-long-term-changes-lmm.png", f)
    f
end