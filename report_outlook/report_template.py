import xlwings as xw
import typing
from report_outlook.component.complex_component import Complex_component
from report_outlook.component.routes_analysis_component import Routes_analysis_component


class Report_template(Complex_component, Routes_analysis_component):
    def __init__(self):
        super().__init__()
        return

    def weekly_op_summary(
            self,
            wb: object,
            ws_name: list,
            date: str,
            endDate:str,
            op_inc: object = {},
            op_exp: object = {},
            op_salary: object = {},
            mv_exp: object = {},
            admin_exp: object = {},
            bins_exp: object = {}):

        super().report_headers(
            wb,
            ws_name,
            date,
            endDate,
            "Weekly Financial Report Summary")

    # Anchor Cell at B6
    # Operating Income
    # Table Headers

        op_tb_header = ["Ton", "Rate per Ton", "% of Total Operating Inc"]

        no_tb_header = []

        super().report_formating(
            wb,
            ws_name)

        # Anchor at B10
        # Operating Income
        if op_inc:
            super().session(
                wb,
                ws_name,
                "Operating Income",
                True,
                op_tb_header,
                op_inc,
                10)
        else:
            print("Operating income items is empty")
            return 0

    # Anchor at B24
    # Operating Expense
        if op_exp:
            super().session(
                wb,
                ws_name,
                "Operating Expense",
                False,
                op_tb_header,
                op_exp,
                24)
        else:
            print("Operating Expense items is empty")
            return 0

        if op_salary:
            super().session(
                wb,
                ws_name,
                "Operating Salary",
                False,
                no_tb_header,
                op_salary,
                36)
        else:
            print("Operating Salary items is empty")
            return 0

        if mv_exp:
            super().session(
                wb,
                ws_name,
                "Motor Vehicle Expense",
                False,
                no_tb_header,
                mv_exp,
                43)
        else:
            print("Motor Vehicle expense items is empty")
            return 0

        if admin_exp:
            super().session(
                wb,
                ws_name,
                "General & Administration",
                False,
                no_tb_header,
                admin_exp,
                59)
        else:
            print("Admin expense items is empty")
            return 0

        if bins_exp:
            super().session(
                wb,
                ws_name,
                "Bins Purchase",
                False,
                no_tb_header,
                bins_exp,
                68)
        else:
            print("Bins expense items is empty")
            return 0

        pass
# ===================================================================================

    def by_rev_type(self,
                    wb: object,
                    ws_name: str,
                    date: str,
                    endDate: str,
                    routes_info: object):

        super().report_formating(
            wb,
            ws_name)

        super().report_headers(
            wb,
            ws_name,
            date,
            endDate,
            "Weekly Financial Report Summary")


        (
            super()
            .income_session(wb, ws_name, routes_info)
            .weight_session(wb, ws_name, routes_info, 9)
            .gross_operating_margin(wb, ws_name, routes_info, 13)
        )

        # route_op_inc
        # {
        # routesName : [1,2,3,4,5]
        # routesInc : [M1, M2, M3, M4]
        # }
