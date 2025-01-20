import json
import sys
import os

from utils import parse_holiday_dates, allocate_hours, verify_allocation_constraints, to_excel_file, get_week_number, weeks_for_year

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Error: Config file name is required as first argument")
        sys.exit(1)
        
    config_file = sys.argv[1]
    file_prefix = None
    if len(sys.argv) > 2:
        file_prefix = sys.argv[2] 
    else:
        file_prefix = f"exports{os.path.sep}"

    with open(config_file, 'r') as f:
        config = json.load(f)

    year = config['year']
    min_week_hours = config['min_week_hours']
    max_week_hours = config['max_week_hours'] 
    average_week_hours = config['average_week_hours']
    average_rolling_week_hours = config['average_rolling_week_hours']
    tracking_rolling_weeks = config['tracking_rolling_weeks']
    max_yearly_overtime = config['max_yearly_overtime']
    yearly_overtime_variance = config['yearly_overtime_variance']
    number_working_days = config['number_working_days']
    project_distribution = config['project_distribution']
    project_names = config['project_names']
    date_format = config['date_format']
    start_date = config['start_date']
    end_date = config['end_date']
    working_days = config['working_days']

    if tracking_rolling_weeks is None:
        tracking_rolling_weeks = 1
        average_rolling_week_hours = max_week_hours

    if number_working_days is None:
        number_working_days = 5

    if date_format is None:
        date_format = "dd/mm"

    start_date = parse_holiday_dates(year, [start_date], format=date_format)[0]
    end_date = parse_holiday_dates(year, [end_date], format=date_format)[0]

    start_week = get_week_number([start_date])[0]
    end_week = get_week_number([end_date])[0]

    assert sum(project_distribution) == 1, "Project distribution must sum to 1"
    assert len(project_distribution) == len(project_names), "Project distribution must have the same length as project names"
    assert min_week_hours <= max_week_hours, "Min week hours must be less than max week hours"
    assert min_week_hours <= average_week_hours <= max_week_hours, "Average week hours must be between min and max week hours"
    assert min_week_hours <= average_rolling_week_hours, "Average rolling week hours must greater than min week hours"
    assert tracking_rolling_weeks > 0, "Tracking rolling weeks must be strictly greater than 0"
    assert max_yearly_overtime >= 0, "Max yearly overtime must be greater than 0"
    assert number_working_days > 0, "Number working days must be strictly greater than 0"
    assert isinstance(year, int), f"Year {year} must be an integer"
    assert yearly_overtime_variance <= max_yearly_overtime, "Min yearly overtime must be less than max yearly overtime"
    assert yearly_overtime_variance >= 0, "Min yearly overtime must be greater than 0"
    assert start_week <= end_week, f"Start week {start_week} must be less than end week {end_week}"
    assert start_week > 0, f"Start week {start_week} must be greater than 0"
    assert end_week <= weeks_for_year(year), f"End week {end_week} must be less than the number of weeks in the year {weeks_for_year(year)}"

    full_holidays = set([h.replace('=f', '') for h in config['holidays'] if ('=' not in h) or h.endswith('=f')])
    mid_holidays = set([h.replace('=m', '') for h in config['holidays'] if h.endswith('=m')]) - full_holidays

    full_holidays = parse_holiday_dates(year, list(full_holidays), format=date_format)
    mid_holidays = parse_holiday_dates(year, list(mid_holidays), format=date_format)

    holiday_dates = {
        "full": full_holidays,
        "mid": mid_holidays
    }

    for label in ["full", "mid"]:
        hd = []
        for d in holiday_dates[label]:
            if start_date <= d  <= end_date:
                hd.append(d)
            else:
                print(f"Excluding holiday {d} outside of start/end date range")
        holiday_dates[label] = hd

    if working_days is not None and len(working_days) > 0:
        working_day_indexes = []
        for d in working_days:
            if d.lower() == "monday":
                working_day_indexes.append(0)
            elif d.lower() == "tuesday":
                working_day_indexes.append(1)
            elif d.lower() == "wednesday":
                working_day_indexes.append(2)
            elif d.lower() == "thursday":
                working_day_indexes.append(3)
            elif d.lower() == "friday":
                working_day_indexes.append(4)
            elif d.lower() == "saturday":
                working_day_indexes.append(5)
            elif d.lower() == "sunday":
                working_day_indexes.append(6)
            else:
                raise ValueError(f"Invalid working day: {d}")
        
        for label in ["full", "mid"]:
            hd = []
            for h in holiday_dates[label]:
                if h.weekday() not in working_day_indexes:
                    print(f"Excluding holiday {h} outside of working day range")
                else:
                    hd.append(h)
            holiday_dates[label] = hd

    print("Start week:", start_week, "End week:", end_week)

    allocation_df = allocate_hours(
            year, holiday_dates, min_week_hours, max_week_hours, average_week_hours,
            average_rolling_week_hours, tracking_rolling_weeks, project_distribution, max_yearly_overtime, yearly_overtime_variance,
            number_working_days=number_working_days, dirichlet_factor=10, start_week=start_week, end_week=end_week
        )

    print("Constraints validation...")

    valid = verify_allocation_constraints(allocation_df, year, holiday_dates, min_week_hours, max_week_hours, 
                                    average_rolling_week_hours, tracking_rolling_weeks,
                                    average_week_hours, max_yearly_overtime, yearly_overtime_variance, number_working_days=number_working_days,
                                    start_week=start_week, end_week=end_week)
    print("Constraints validation result:", valid)

    assert valid, "Allocation is not valid"

    filename = to_excel_file(allocation_df, project_names,
                title=f"Allocation du temps {year} (en heures)",
                week_col_prefix="S",
                project_col_title="Tâches",
                file_prefix=file_prefix
            )

    print(f"Allocation saved to {filename}")
