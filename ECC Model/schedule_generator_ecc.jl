using JuMP,DataFrames,XLSX,Dates,Plots
import HiGHS

# Helper Functions: 

function fill_hours(start=1, stop=24)
    day_avail = [j in start:stop ? 1 : 0 for j in 1:24]
    return day_avail
end

function fill_week(df_req)
    num_units = size(df_req)[1]
    unit_open = zeros(24, 7,num_units)
    for n in 1:num_units
        hours_open_start = df_req.open_start[n]
        hours_open_end = df_req.open_end[n]

        for i in 1:7
            unit_open[:, i, n] .= fill_hours(hours_open_start, hours_open_end)
        end
    end
    return unit_open
end

# Group the indices of the changes by consecutive values, useful when someone works a night/morning shift
function groups_mult(binary_array)
    n_index = 1
    for n in 1:num_areas
        if sum(binary_array[:,n]) > 0
            n_index = n
        end
    end
    groups = [[],[],[],[]]
    current_group = []
    current_index = 1
    for i = 1:size(binary_array)[1]
        if sum(binary_array[i,:]) > 0
            push!(current_group, i)
        elseif length(current_group) > 0
            groups[current_index] = current_group
            current_group = []
            current_index += 1
        end
    end
    if length(current_group) > 0
        groups[current_index] = current_group

    end
    return groups, n_index
end

function generate_schedule(Y_results,data)
    I_len = length(data.name)
    A = fill("X", I_len,9)
    A[:,1] = data.name
    A[:,2] = [string(sum(Y_results[i,:,:,:,:])) for i in 1:I_len]
    df_cols = ["Name","Scheduled Hours","Sun","Mon","Tue","Wed","Thu","Fri","Sat"]

    for i in 1:I_len,d in 1:7, role_i in R
        Y_index = Y_results[i,:,d,:,role_i]
        day_hours = Matrix(Y_index)
        if sum(day_hours) >= 0.5
            grouped = groups_mult(Y_index)
            groups = grouped[1]
            if length(groups[2]) > 1
                n_index = grouped[2]
                shift1 = groups[1]
                shift1_start,shift1_end = shift1[1], shift1[length(shift1)]
                start_1_am_pm = Dates.format(Dates.Time(shift1_start-1), dateformat"II:MM p")
                end_1_am_pm = Dates.format(Dates.Time(shift1_end-1)+Dates.Hour(1), dateformat"II:MM p")

                shift2 = groups[2]
                shift2_start,shift2_end = shift2[1], shift2[length(shift2)]
                start_2_am_pm = Dates.format(Dates.Time(shift2_start-1), dateformat"II:MM p")
                end_2_am_pm = Dates.format(Dates.Time(shift2_end-1)+Dates.Hour(1), dateformat"II:MM p")
                assignment = string(uppercase(roles[role_i]),
                    " in ",areas[n_index],": ",
                    start_1_am_pm,"-",end_1_am_pm," and ", start_2_am_pm,"-",end_2_am_pm)
                A[i,d+2] = assignment

            else
                shift_start,area_n = Tuple(findfirst(!iszero,day_hours))
                start_am_pm = Dates.format(Dates.Time(shift_start-1), dateformat"II:MM p")
                shift_end =findlast(!iszero,day_hours)[1]
                end_am_pm = Dates.format(Dates.Time(shift_end-1)+Dates.Hour(1), dateformat"II:MM p")

                assignment = string(uppercase(roles[role_i])," in ",areas[area_n],": ",start_am_pm,"-",end_am_pm)

                A[i,d+2] = assignment
            end
        end
    end

    df = DataFrame(A,df_cols)
    return df
end

### Reading in data 

filename = "staff_info_ecc.xlsx"
staff_pref = DataFrame(XLSX.readtable(filename, "staff_preferences"))
df1 = DataFrame(XLSX.readtable(filename, "week1"))
df2 = DataFrame(XLSX.readtable(filename, "week2"))
df_req = DataFrame(XLSX.readtable(filename, "requirements"))

num_staff = length(staff_pref.name)

areas = df_req.areas
num_areas = length(areas)

min_req = df_req[!,r"min"] # min demand in each area (n,r) --> (CCU:ER, CVT:AST)
max_req = df_req[!,r"max"] # max demand in each area (n,r) --> (CCU:ER, CVT:AST)

df_areas = (staff_pref[!,r"area"]) # area for each employee (I, CCU:ER)
df_roles = staff_pref[:,r"role"] # roles for each employee (I, CVT,AST,CCL)
df_roles.role_assistant = ifelse.(df_roles.role_cvt .== 1, 0,df_roles.role_assistant)
night_pref = staff_pref.night_pref



# Establishing sets and parameters

global I = 1:num_staff # list of receptionist index
shifts = 1:24 # list of shift index
days = 1:7 # list of day index
N = 1:num_areas # list of areas index

max_weekly_hour = staff_pref.max_hours # maximum hours for each worker
min_weekly_hour = staff_pref.min_hours # minimum hours for each worker
hour_allowance = max_weekly_hour-min_weekly_hour

min_dur = staff_pref.min_shift_length # min shift duration
max_dur = staff_pref.max_shift_length # max shift duration



hours_open_start = df_req.open_start # 
hours_open_end = df_req.open_end

night_shift_start = df_req.night_shift_start
night_shift_end = df_req.night_shift_end


week_open = fill_week(df_req)


R = 1:size(df_roles)[2]
roles = ["cvt","ast","ccl"]
night_shift_pref = ([[j in 1:night_shift_end[1] || j in night_shift_start[1]:24 ? 10*night_pref[i]+1 : 1 for j in 1:24] for i in I])


function optimize_schedule(data)
    # Generating employee availability arrays
    df_avail_start = data[!,r"start"]
    df_avail_end = data[!,r"end"]
    num_staff, cols = size(data)
    
    employee_avail = zeros(num_staff, 24,7)

    for i in I
        for d in 1:7
            start = df_avail_start[i,d]
            hour_end = df_avail_end[i,d]
            x = fill_hours(start,hour_end)
            employee_avail[i,:,d] = x
        end
    end
    
    ##### Create model ######
    m = Model()

    # Create decision variables
    @variable(m, Y[I, shifts, days, N,R], Bin) # Staff assignment Y[staff,hour,day,area, role]
    @variable(m, X[I, days, N,R], Bin) # Day working assignment, a 1 indicates the day/desk a emplyee is working
    @variable(m, S[I, shifts, days, N,R], Bin) # "Start-ups" variable, a 1 indicates the day/time/desk a employee starts working
    @variable(m, P[shifts, days, N,R] >= 0, Int) # Additional staff necessary

    # Set objective
    @objective(m, Max, 
        sum(df_roles[i,r]*df_areas[i,n]*employee_avail[i,j,d]*Y[i, j, d, n,r]*night_shift_pref[i][j] 
            for i in I, j in shifts, d in days, n in N, r in R))
    #+ 0.01*sum(P[j,d,n,r] for j in shifts, d in days, n in N, r in R))


    # Setting constraints
    for i in I

        # Staff can only be scheduled to work between their min and max total hours
        @constraint(m, 0 <= max_weekly_hour[i] - 
            sum(Y[i,j,d,n,r] for j in shifts, d in days, n in N, r in R) <= hour_allowance[i])


        # Each person only works one weekend shift
        @constraint(m, sum(X[i,1,n,r] + X[i,7,n,r] for n in N for r in R) <= 1) # Change this to allow more weekend shifts
        
        # Each person can only work 5 days per week
        @constraint(m, sum(X[i,d,n,r] for d in days for n in N for r in R) <= 5) # Change this to allow more days per week


        for d1 in 1:6, n in N, r in R

            # Ensuring the amount of hours worked during night shift are within bounds            
            @constraint(m, 
                sum(Y[i,j1, d1, n,r] for j1 in night_shift_start[n]:24) +
                sum(Y[i,j2, d1+1, n,r] for j2 in 1:night_shift_end[n]) <= max_dur[i])

            for j1 in night_shift_start[n]:24
                # If a worker starts a night shift they then work the next morning also
                @constraint(m, S[i,j1,d1,n,r] == sum(S[i,1,d1+1,n1,r1] for n1 in N,r1 in R))
            end
        end

        for d in days
            # Each person can only work one area per day on the day they are available
            @constraint(m, sum(X[i, d, n, r] for n in N for r in R) <= 1)        


            for n in N
                for r in R
                    # Constraining minimum and maximum daily hours in a single area and role
                    @constraint(m,sum(Y[i,j, d, n,r] for j in shifts) <= max_dur[i]*X[i,d,n,r])
                    @constraint(m,min_dur[i]*X[i,d,n,r] <= sum(Y[i,j, d, n,r] for j in shifts))

                    # Defines the start-up variable in relation to the Y variable
                    for j in 2:24   
                        @constraint(m, S[i,j,d,n,r] >= Y[i, j, d, n,r]- Y[i,j-1,d,n,r])
                    end
                end

                # Limits Y variable to one "start-up" such that every employee works consecutive hours  
                @constraint(m,sum(S[i,j,d,n,r] for j in shifts for r in R) <= 1)

            end
        end
    end


    # Add constraint settting minimum between nightshift and end of nightshift for ER

    for j in shifts, d in days, n in N, r in R

        # Each area must be filled during open hours
        @constraint(m, min_req[n,r]*week_open[j,d,n]<=
            sum(Y[i,j,d,n,r] for i in I) + P[j,d,n,r] <=  max_req[n,r]*week_open[j,d,n])
    end

    for n in N,j1 in hours_open_start[n]:hours_open_end[n], d in days, r in R
    # Ensure at least one person works in each area during open hours
        @constraint(m, sum(Y[i,j1,d,n,r] for i in I) >= min_req[n,r])
    end

    set_optimizer(m, HiGHS.Optimizer,add_bridges = false)
    set_optimizer_attribute(m, "presolve", "on")
    set_optimizer_attribute(m, "mip_rel_gap", 0.01)
    set_optimizer_attribute(m, "time_limit", 1000.0)
    optimize!(m)

    Y_val = value.(Y) # Our staff variable
    X_val = value.(X) # Our day assignment variable
    S_val = value.(S) # Startup variable
    P_val = value.(P) # Slack variable
    solution_summary(m,verbose=true)
    println(sum(Y_val),",",sum(P_val))

    schedule = generate_schedule(Y_val,data) # Creating week i's schedule
    variable_result = Dict("Y"=>Y_val,"X"=>X_val,"S"=>S_val,"P"=>P_val,"schedule"=>schedule) # Saving model output into a dict
    return variable_result
end

### Getting Results

data_list = [df1,df2]
results_list = [optimize_schedule(i) for i in data_list]

today_filename = Date(now())
today = Date(Dates.firstdayofweek(now())-Day(1))
nextweek = today+Day(7)
filename = "schedule_ecc_$today_filename.xlsx"

sheet_names = [ "Schedule A $today","Schedule B $nextweek"]
dataframes = [results_list[i]["schedule"] for i in 1:2]

# Writing created schedules into Excel, each week get its own sheet
XLSX.writetable(filename, overwrite=true, 
    sheet_names[1] => dataframes[1],
    sheet_names[2] => dataframes[2])

#covered = reshape(sum(Y_results[i,:,:,:,:] for i in 1:I_len),(24,7*2*2))
#covered_cols = [i*"_"*j for j in roles for i in df_cols[3:9]]
#df_covered = DataFrame(covered,covered_cols)



using Plots.PlotMeasures

wday = ["Sun" "Mon" "Tue" "Wed" "Thu" "Fri" "Sat"]
unworked = [string(round(Int,sum(results_list[z]["P"]))) for z in 1:2]

role_num = length(roles)
role_labs  = reshape(roles,(1,role_num))

covered_cols = Matrix(reshape([j*"_"*l for j in areas,l in roles],(1,role_num*num_areas)))

results = [sum(results_list[z]["Y"][i,:,:,:,:] for i in I) for z in 1:2]
results_vectors = [[[results[i][:,d,:,1] results[i][:,d,:,2]] for d in 1:7] for i in 1:2] # (CCU,CVT) (ER,CVT) (CCU, AST) (ER, AST)

plt = [[plot(results_vectors[i][d], title= " Week $i: Coverage for "*wday[d],ylim=(0,5),labels=covered_cols,
        xticks=1:24,legend=false, line=(2.5,:dash)) for d in 1:7] for i in 1:2]

pL = [plot((1:role_num*num_areas)', title="Total Hours Unscheduled Week $z: "*unworked[z],
    legend=:bottom, legendcolumns=role_num*num_areas,labels=covered_cols,framestyle=:none) for z in 1:2]

pzz = [pL[1],plt[1]...,pL[2],plt[2]...]
pz = plot(pzz..., layout = grid(7*2+2,1), size = (1600, 4550),left_margin=15mm,top_margin=15mm)
savefig(pz, "ecc_coverage_plots.pdf") # save the most recent fig as filename_string (such as "output.png")