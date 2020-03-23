"""
    CopyRight Dr. Ahmad Hamdi Emara 2020
    Adam Medical Company
"""

import sys

def main():
    if len(sys.argv) > 1:
        salary(True)
    else:
        salary(False)


def salary(args_availble):
    default_list = 880.0
    default_overtime = 35.0
    default_discount = 974.0
    default_sales = 158600/2.0
    default_bonus = 800.0

    base_salary = 4500.0
    print("=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=")

    print(f"Base salary: {base_salary}")
    single_housing_allowance = round(15000.0 / 12.0, 2)
    print(f"Single housing allowance:  {single_housing_allowance}")
    married_housing_allowance = round(17000.0 / +12.0, 2)
    print(f"Married housing allowance: {married_housing_allowance}")
    

    
    overtime_hrs = default_overtime
    if args_availble:
        overtime_hrs = float(sys.argv[1])

    print(f"Overtime hours: {overtime_hrs}")


    lst = float(sys.argv[2]) or default_list if args_availble else default_list
    sold_list = round((100.0 * lst) / 7.0, 2)

    
    total_sales = round((float(sys.argv[3]) or default_sales) - sold_list if args_availble else default_sales - sold_list, 2)
    print(f"Total sales without the list items: {total_sales}" )
    print("=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=")

    cut = round(total_sales * 0.01, 2)
    print(f"Sales cut: {cut}")
    print("=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=")

    discounted = (float(sys.argv[4]) or default_discount) if args_availble else default_discount

    overtime_hr_value = round((base_salary / 26.0 / 10.0) * 1.5, 2)
    total_overtime = round(overtime_hrs * overtime_hr_value, 2)
    print(f"Total overtime value: {total_overtime}")
    print("=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=")
    
    print(f"This month bonuses {default_bonus}")
    print("=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=")

    total_salary_without_allowance = round(base_salary - discounted + total_overtime + lst + cut + default_bonus, 2)
    total_married_salary = round(total_salary_without_allowance + married_housing_allowance, 2)
    total_single_salary = round(total_salary_without_allowance + single_housing_allowance, 2)
    
    print("=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=")

    print(f"Married salary: {total_married_salary}")

    print("=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=")

    print(f"Single salary: {total_single_salary}" )

    print("=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=")

if __name__ == "__main__":
    main()