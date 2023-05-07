using JuMP,DataFrames,XLSX,Dates,Plots
import HiGHS


# Importing data from Excel
filename = "staff_info_reception.xlsx"
staff_pref = DataFrame(XLSX.readtable(filename, "staff_preferences"))
df1 = DataFrame(XLSX.readtable(filename, "week1"))
df2 = DataFrame(XLSX.readtable(filename, "week2"))
hours_df = DataFrame(XLSX.readtable(filename, "requirements"))
desk_pref = Matrix(staff_pref[!,r"pref"])
num_desks = size(hours_df)[1]

# Error handling incase desks are added incorrectly
I, desks_prefered = size(desk_pref)
if desks_prefered != num_desks
    d_num = abs(desks_prefered - num_desks)
    desk_pref = [desk_pref Int.(ones(size(desk_pref)[1],d_num))]
end


### Helper Functions: 

# Encodes 24hr availability given a start and stop time
function fill_hours(start=1, stop=24)
    day_avail = [j in trunc(Int, start):trunc(Int,stop) ? 1 : 0 for j in 1:24]
    return day_avail
end

# Encodes 24hr/7day availability given a start and stop time
function fill_week(hours_data)
    num_desks = size(hours_data)[1]
    desk_open = zeros(24, 7,num_desks)
    for j in 1:num_desks
        week_open = hours_data.week_open[j]
        week_close = hours_data.week_close[j]
        weekend_open = hours_data.weekend_open[j]
        weekend_close = hours_data.weekend_close[j]

        for i in 2:6
            desk_open[:, i, j] .= fill_hours(week_open, week_close)
        end

        for i in [1, 7]
            desk_open[:, i, j] .= fill_hours(weekend_open, weekend_close)
        end
    end
    return desk_open
end

# Creates a dataframe from the model results including total hours scheduled, desk assignment and hours in AM/PM
function generate_schedule(Y_results, data)
    I_len = length(data.name)
    A = fill("X", I_len,9)
    A[:,1] = data.name
    A[:,2] = [string(round(Int,(sum(Y_results[i,:,:,:])))) for i in 1:I_len]
    df_cols = ["Receptionist","Scheduled Hours","Sun","Mon","Tue","Wed","Thu","Fri","Sat"]

    for i in 1:I_len,d in 1:7
        day_hours = Matrix(Y_results[i,:,d,:])
        if sum(day_hours) >= 0.5
            shift_start,desk_n = Tuple(findfirst(!iszero,day_hours))
            start_am_pm = Dates.format(Dates.Time(shift_start-1), dateformat"II:MM p")
            shift_end =findlast(!iszero,day_hours)[1]
            end_am_pm = Dates.format(Dates.Time(shift_end-1)+Dates.Hour(1), dateformat"II:MM p")
            
            assignment = string(desks[desk_n],":",start_am_pm,"-",end_am_pm)

            A[i,d+2] = assignment
        end
    end

    df = DataFrame(A,df_cols)
    
    covered = reshape(sum(Y_results[i,:,:,:] for i in 1:I_len),(24,7*num_desks))
    covered_cols = [i*"_"*j for j in desks for i in df_cols[3:9]]
    df_covered = DataFrame(covered,covered_cols)
    
    return df,df_covered
end

desk_open = Int.(fill_week(hours_df)) # creates a binary weeklong schedule for each desk (hour,day, desk)
desks = hours_df.desk # list of desk names for assignment
num_desks = length(desks) # total number of desks

min_weekly_hour = staff_pref.min_hours # minimum hours for each receptionist
max_weekly_hour = staff_pref.max_hours # maximum hours for each 

min_req = hours_df.min_staff # min demand at each desk
max_req = hours_df.max_staff # max demand at each desk

min_dur = staff_pref.min_shift_length # min shift duration
max_dur = staff_pref.max_shift_length # max shift duration



### Building our model

function optimize_schedule(data)
    # Generating employee availability arrays
    df_avail_start = data[!,r"start"]
    df_avail_end = data[!,r"end"]

    num_staff, cols = size(data)
    
    
    # Establishing sets and parameters
    
    global I = 1:num_staff # list of receptionist index
    shifts = 1:24 # list of shift index
    days = 1:7 # list of day index
    N = 1:num_desks # list of desk index
    
    employee_avail = zeros(num_staff, 24,7)

    for i in I
        for d in 1:7
            start = df_avail_start[i,d]
            hour_end = df_avail_end[i,d]
            x = fill_hours(start,hour_end)
            employee_avail[i,:,d] = x
        end
    end


    # Create model
    m = Model()

    # Create decision variables
    @variable(m, Y[I, shifts, days, N], Bin) # Staff assignment
    @variable(m, X[I, days, N], Bin) # Day working assignment, a 1 indicates the day/desk a emplyee is working
    @variable(m, S[I, shifts, days, N], Bin) # "Start-ups" variable, a 1 indicates the day/time/desk a employee starts working
    @variable(m, P[shifts, days, N] >= 0, Int) # Additional staff necessary

    # Set objective
    @objective(m, Max, 
        sum(desk_pref[i,n]*employee_avail[i,j,d]*Y[i, j, d, n] for i in I, j in shifts, d in days, n in N))


    # Setting constraints
  
    for i in I

        # Each person can only work 5 days per week
        @constraint(m, sum(X[i,d,n] for d in days for n in N) <= 5)

        for d in days

            # Each person can only work one desk per day on the day they are available
            @constraint(m, sum(X[i, d, n] for n in N) <= 1)        

            for n in N

                # Constraining minimum and maximum daily hours at a single desk
                @constraint(m,sum(Y[i,j, d, n] for j in shifts) <= max_dur[i]*X[i,d,n])
                @constraint(m,min_dur[i]*X[i,d,n] <= sum(Y[i,j, d, n] for j in shifts))


                # Defines the start-up variable in relation to the Y variable
                for j in 2:24   
                    @constraint(m, S[i,j,d,n] >= Y[i, j, d, n]- Y[i,j-1,d,n])
                end

                # Limits Y variable to one "start-up" such that every employee works consecutive hours  
                @constraint(m,sum(S[i,j,d,n] for j in shifts) <= 1)
            end

        end

        # Staff can only be scheduled to work between their min and max total hours
        @constraint(m, min_weekly_hour[i] <= sum(Y[i,j,d,n] for j in shifts for d in days for n in N) <= max_weekly_hour[i])


        # Each person only works one weekend shift
        @constraint(m, sum(X[i,1,n] + X[i,7,n] for n in N) <= 1)

    end

    for j in shifts, d in days, n in N

        # Each desk must be filled during open hours
        @constraint(m, desk_open[j,d,n]*min_req[n] 
                        <= sum(Y[i,j,d,n] for i in I) + P[j,d,n] <=  
                       desk_open[j,d,n]*max_req[n])

        @constraint(m, sum(Y[i,j,d,n] for i in I) >= desk_open[j,d,n])
        
    end

    set_optimizer(m, HiGHS.Optimizer)
    set_optimizer_attribute(m, "presolve", "on")
    set_optimizer_attribute(m, "time_limit", 1000.0)
    optimize!(m)

    Y_val = value.(Y) # Our staff variable
    X_val = value.(X) # Our day assignment variable
    S_val = value.(S) # Startup variable
    P_val = value.(P) # Slack variable
    solution_summary(m,verbose=true)
    schedule,coverage = generate_schedule(Y_val,data) # Creating week i's schedule
    variable_result = Dict("Y"=>Y_val,"X"=>X_val,"S"=>S_val,"P"=>P_val,"schedule"=>schedule,"coverage"=>coverage) # Saving model output into a dict
    return variable_result
end


### Getting Results

data_list = [df1,df2]
results_list = [optimize_schedule(i) for i in data_list]

today_filename = Date(now())
today = Date(Dates.firstdayofweek(now())-Day(1))
nextweek = today+Day(7)
filename = "schedule_reception_$today_filename.xlsx"

sheet_names = [ "Schedule A $today","Schedule B $nextweek"]
dataframes = [results_list[i]["schedule"] for i in 1:2]

# Writing created schedules into Excel, each week get its own sheet
XLSX.writetable(filename, overwrite=true, 
    sheet_names[1] => dataframes[1],
    sheet_names[2] => dataframes[2])


# Plotting how many receptionists are at each desk throughout each week

using Plots.PlotMeasures

unworked = [string(floor(sum(results_list[j]["P"]))) for j in 1:2]
x = [results_list[d]["coverage"][:,[(1+7*(i-1)+k) for i in 1:num_desks]] for d in 1:2 for k in 0:6]

deskz = reshape(desks,(1,num_desks))

wday = ["Sun" "Mon" "Tue" "Wed" "Thu" "Fri" "Sat"]

p = [plot(Matrix(x[i]),ylim=(0,5),
        legend=false,xticks=1:24,xlab="Hour", title="Week $w: Coverage for "*wday[i],
        line=(2,:dash)) for w in 1:2, i in 1:7]

weeks = ["Week1", "Week2"]
# legend plot
pL = [plot((1:num_desks)', title="Total Hours Unscheduled Week $w: "*unworked[w],
    legend=:inside, legendcolumns=num_desks,line=(2.5,:dash),
    framestyle=:none,labels=deskz) for w in 1:2]

p1 = p[1,:]
p2 = p[2,:]
pzz = [pL[1],p1...,pL[2],p2...]

pz = plot(pzz..., layout = grid(7*2+2,1), size = (1600, 4550),left_margin=10mm,top_margin=15mm)
savefig(pz, "reception_coverage_plots.pdf") # save the most recent fig as filename_string (such as "output.png")

