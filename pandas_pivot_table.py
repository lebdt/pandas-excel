import datetime
import pandas as pd
import numpy as np
import xlsxwriter.utility as xu

np.random.seed(42)

def ExampleOne():
    x = pd.DataFrame(
        {
            "G": [np.random.random() for _ in range(6)],
            # "test": [],
            # "value" : [],
            "var": ["B", "B", "B", "C", "C", "C"],
            "date": [
                "2024-01-01",
                "2024-02-01",
                "2024-03-01",
                "2024-01-01",
                "2024-02-01",
                "2024-03-01",
            ],
        }
    )

    y = pd.DataFrame(
        {
            "H": [np.random.random(), np.random.random(), np.random.random()],
            # "test": [],
            # "value" : [],
            "var": ["A", "B", "B"],
            "date": ["2024-01-01", "2024-02-01", "2024-03-01"],
        }
    )


    x["value"] = x["G"]
    x["test"] = "G"
    x = x.drop("G", axis=1)


    y["value"] = y["H"]
    y["test"] = "H"
    y = y.drop("H", axis=1)


    # print("x\n %s\n" % x)
    # print(y)


    outer = x.merge(y, on=["var", "date", "value", "test"], how="outer")
    pt = pd.pivot_table(outer, columns="date", index=["var", "test"], values="value")


    def change_date_cols(df):
        c_df = pd.DataFrame(df)
        c_df.columns = [
            datetime.datetime.strptime(col, r"%Y-%m-%d").strftime(r"%b-%y")
            for col in c_df.columns
        ]
        return c_df


    df = change_date_cols(pt)


    def save_to_excel(df, sheet_name="Sheet1", st_row=5, st_col=0):
        xl = "/Users/ebd/pandas_pivot_table.xlsx"
        with pd.ExcelWriter(path=xl, engine="xlsxwriter") as writer:
            df.to_excel(
                writer,
                sheet_name=sheet_name,
                freeze_panes=(st_row + 1, st_col + len(df.index[0])),
                startrow=st_row,
                startcol=st_col,
            )

            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            worksheet.insert_image(
                "A1",
                "example.png",
                {"x_offset": 15, "y_offset": 10, "x_scale": 0.35, "y_scale": 0.35},
            )

            worksheet.hide_gridlines(2)

            header_format = workbook.add_format(
                {
                    "bg_color": "#dd2119",
                    "font_color": "#fefefe",
                    "border": 1,
                    "border_color": "#fefefe",
                    "bold": True,
                }
            )

            special_index_format = workbook.add_format(
                {
                    "bg_color": "#cbcbcb",
                    "font_color": "#dd2119",
                    "border": 1,
                    "border_color": "#dd2119",
                    "bold": True,
                }
            )

            outer_format = workbook.add_format(
                {"top": 3, "left": 3, "bottom": 3, "right": 3}
            )

            def color_header(
                df, ws=worksheet, f=header_format, st_row=st_row, st_col=st_col
            ):
                header_range = xu.xl_range(
                    st_row, st_col, st_row, st_col + len(df.columns) + len(df.index[0]) - 1
                )
                ws.conditional_format(header_range, {"type": "no_errors", "format": f})
                return

            def color_indices(
                df,
                ws=worksheet,
                f=header_format,
                si=special_index_format,
                st_row=st_row,
                st_col=st_col,
            ):
                index_range = xu.xl_range(
                    st_row, st_col, st_row + len(df), st_col + len(df.index[0]) - 2
                )
                ws.conditional_format(index_range, {"type": "no_errors", "format": f})
                si_pos = len(df.index[0][:-1])
                si_range = xu.xl_range(
                    st_row + 1, st_col + si_pos, st_row + len(df), st_col + si_pos
                )
                ws.conditional_format(si_range, {"type": "no_errors", "format": si})
                return

            def get_inside_tab_range(df):
                top = 1
                res = xu.xl_range(
                    st_row + top,
                    st_col + len(df.index[0]),
                    st_row + len(df),
                    st_col + len(df.columns) + len(df.index[0]) - 1,
                )
                return res

            def get_outer_range(df):
                res = xu.xl_range(
                    st_row,
                    st_col,
                    st_row + len(df),
                    st_col + len(df.columns) + len(df.index[0]) - 1,
                )
                return res


            color_header(df)
            color_indices(df)
            # x = xu.xl_range(0, 0, pt.shape[0], pt.shape[1]+len(pt.index[0])-1)
            # worksheet.conditional_format(x, {'type': 'no_errors', 'format': header_format})

            inner_format = workbook.add_format(
                {"num_format": "0.00%", "border": 1, "border_color": "#dd2119"}
            )

            alignment = workbook.add_format({"valign": "vcenter", "align": "center"})

            # worksheet.conditional_format(get_outer_range(df), {"type": "no_errors", "format" :outer_format})
            worksheet.conditional_format(
                get_inside_tab_range(df), {"type": "no_errors", "format": inner_format}
            )

            full_range = xu.xl_range(
                st_row,
                st_col,
                st_row + len(df),
                st_col + len(df.columns) + len(df.index[0]) - 1,
            )
            worksheet.set_column(full_range, None, alignment)

            def get_indices_for_value(df, value):
                index_col_dlist = [[] for _ in range(len(df.index[0]))]
                for j in range(len(df.index)):
                    count = 0
                    for elem in df.index[j]:
                        if value == elem:
                            index_col_dlist[count].append(j + 1)
                        count += 1
                return index_col_dlist

            # H_row = 4
            red_format = workbook.add_format({"bg_color": "#ff9090"})
            green_format = workbook.add_format({"bg_color": "#90ff90"})
            yellow_format = workbook.add_format({"bg_color": "#ffff90"})
            gray_format = workbook.add_format({"bg_color": "#909090"})

            G_indices = get_indices_for_value(df, "G")[1]
            for pos in G_indices:
                G_range = xu.xl_range(
                    st_row + pos,
                    st_col + len(df.index[0]),
                    st_row + pos,
                    st_col + len(df.columns) + len(df.index[0]) - 1,
                )
                worksheet.conditional_format(
                    G_range,
                    {
                        "type": "cell",
                        "criteria": "equal to",
                        "value": '"-"',
                        "format": gray_format,
                    },
                )
                worksheet.conditional_format(
                    G_range,
                    {
                        "type": "cell",
                        "criteria": "greater than or equal to",
                        "value": 0.5,
                        "format": green_format,
                    },
                )
                worksheet.conditional_format(
                    G_range,
                    {
                        "type": "cell",
                        "criteria": "less than",
                        "value": 0.3,
                        "format": red_format,
                    },
                )
                worksheet.conditional_format(
                    G_range,
                    {
                        "type": "cell",
                        "criteria": "between",
                        "minimum": 0.3,
                        "maximum": 0.5,
                        "format": yellow_format,
                    },
                )

            H_indices = get_indices_for_value(df, "H")[1]
            for pos in H_indices:
                H_range = xu.xl_range(
                    st_row + pos,
                    st_col + len(df.index[0]),
                    st_row + pos,
                    st_col + len(df.columns) + len(df.index[0]) - 1,
                )
                worksheet.conditional_format(
                    H_range,
                    {
                        "type": "cell",
                        "criteria": "equal to",
                        "value": '"-"',
                        "format": gray_format,
                    },
                )
                worksheet.conditional_format(
                    H_range,
                    {
                        "type": "cell",
                        "criteria": "less than",
                        "value": 0.25,
                        "format": green_format,
                    },
                )
                worksheet.conditional_format(
                    H_range,
                    {
                        "type": "cell",
                        "criteria": "greater than or equal to",
                        "value": 0.25,
                        "format": red_format,
                    },
                )
        return


    df_fillna = df.fillna("-")
    save_to_excel(df_fillna, st_col=2)

    print(df_fillna)
    return


def ExampleTwo():
    x = pd.DataFrame(
        {
            "L": [np.random.random() for _ in range(6)],
            # "test": [],
            # "value" : [],
            "var": ["B", "B", "B", "C", "C", "C"],
            "SEG": ["S1", "S2", "S3", "S1", "S2", "Missing"],
            "#": ["1", "2", "3", "1", "2", "3"]
        }
    )

    y = pd.DataFrame(
        {
            "L": [np.random.random() for _ in range(4)],
            # "test": [],
            # "value" : [],
            "var": ["A", "A", "D", "D"],
            "SEG": ["S1", "S2", "S1", "S2"],
            "#": ["1", "2", "1", "2"]
        }
    )


    x["value"] = x["L"]
    x = x.drop("L", axis=1)


    y["value"] = y["L"]
    y = y.drop("L", axis=1)


    # print("x\n %s\n" % x)
    # print(y)


    outer = x.merge(y, on=["var", "value", "SEG", "#"], how="outer")
    indexed_df = outer.set_index(["var", "#", "SEG"]).sort_index()

    df = indexed_df

    df_test = pd.DataFrame(indexed_df)
    print(df_test.groupby("var")["value"].count().to_dict())

    df_test.drop("value",axis=1, inplace=True)

    for i in range(1,6):
        df_test[str(i)] = np.nan

    for i in range(len(df_test)):
        j = np.random.randint(1,4)
        df_test.iloc[i,j] = np.random.random()

    df_test = df_test.reset_index().drop(["var", "SEG"], axis=1)#.fillna("-")


    def save_to_excel(df, sheet_name="Sheet1", st_row=5, st_col=2):
        xl = "/Users/ebd/pandas_pivot_table_2.xlsx"
        with pd.ExcelWriter(path=xl, engine="xlsxwriter") as writer:
            df.to_excel(
                writer,
                sheet_name=sheet_name,
                freeze_panes=(st_row + 1, st_col + len(df.index[0])),
                startrow=st_row,
                startcol=st_col,
            )

            df_test_st_pos = (st_row,st_col+len(df.columns)+len(df.index[0])+1)

            print(df_test_st_pos)
            print(df.columns)
            df_test.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=df_test_st_pos[0],
                startcol=df_test_st_pos[1],
                index=False
            )

            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            worksheet.insert_image(
                "A1",
                "example.png",
                {"x_offset": 15, "y_offset": 10, "x_scale": 0.35, "y_scale": 0.35},
            )

            worksheet.hide_gridlines(2)

            header_format = workbook.add_format(
                {
                    "bg_color": "#dd2119",
                    "font_color": "#fefefe",
                    "border": 1,
                    "border_color": "#fefefe",
                    "bold": True,
                }
            )

            special_index_format = workbook.add_format(
                {
                    "bg_color": "#cbcbcb",
                    "font_color": "#dd2119",
                    "border": 1,
                    "border_color": "#dd2119",
                    "bold": True,
                }
            )

            outer_format = workbook.add_format(
                {"top": 3, "left": 3, "bottom": 3, "right": 3}
            )

            def color_header(
                df, ws=worksheet, f=header_format, st_row=st_row, st_col=st_col
            ):
                if type(df.index[0]) == tuple:
                    header_range = xu.xl_range(
                        st_row, st_col, st_row, st_col + len(df.columns) + len(df.index[0]) - 1
                    )
                else:
                    header_range = xu.xl_range(
                        st_row, st_col, st_row, st_col + len(df.columns) - 1
                    )

                ws.conditional_format(header_range, {"type": "no_errors", "format": f})
                return

            def color_indices(
                df,
                ws=worksheet,
                f=header_format,
                si=special_index_format,
                st_row=st_row,
                st_col=st_col,
            ):
                index_range = xu.xl_range(
                    st_row, st_col, st_row + len(df), st_col + len(df.index[0]) - 2
                )
                ws.conditional_format(index_range, {"type": "no_errors", "format": f})
                si_pos = len(df.index[0][:-1])
                si_range = xu.xl_range(
                    st_row + 1, st_col + si_pos, st_row + len(df), st_col + si_pos
                )
                ws.conditional_format(si_range, {"type": "no_errors", "format": si})
                return

            def color_col(
                df,
                col,
                ws=worksheet,
                f=special_index_format,
                st_row=st_row,
                st_col=st_col
            ):
                col_num = df.columns.get_loc(col)
                if type(df.index[0]) == tuple:
                    col_num = col_num + len(df.index[0])
                col_range = xu.xl_range(
                    st_row+1, st_col + col_num, st_row + len(df), st_col + col_num
                )
                ws.conditional_format(col_range, {"type": "no_errors", "format": f})
                return

            def get_inside_tab_range(df, st_row=st_row, st_col=st_col):
                top = 1

                col_pos = st_col
                if type(df.index[0]) == tuple:
                    col_pos = st_col + len(df.index[0])

                res = xu.xl_range(
                    st_row + top,
                    col_pos,
                    st_row + len(df),
                    col_pos + len(df.columns) - 1,
                )
                return res

            def get_outer_range(df):
                col_pos = st_col
                if type(df.index[0]) == tuple:
                    col_pos = st_col + len(df.index[0])

                res = xu.xl_range(
                    st_row,
                    st_col,
                    st_row + len(df),
                    col_pos + len(df.columns) - 1,
                )
                return res


            color_header(df)
            color_indices(df)

            color_header(df_test, st_row=df_test_st_pos[0], st_col=df_test_st_pos[1])
            color_col(df_test, "#", st_row=df_test_st_pos[0], st_col=df_test_st_pos[1])
            # color_indices(df_test, st_row=df_test_st_pos[0], st_col=df_test_st_pos[1])

            # x = xu.xl_range(0, 0, pt.shape[0], pt.shape[1]+len(pt.index[0])-1)
            # worksheet.conditional_format(x, {'type': 'no_errors', 'format': header_format})

            inner_format = workbook.add_format(
                {"num_format": "0.00%", "border": 1, "border_color": "#dd2119"}
            )

            alignment = workbook.add_format({"valign": "vcenter", "align": "center"})

            worksheet.conditional_format(
                get_inside_tab_range(df), {"type": "no_errors", "format": inner_format}
            )

            worksheet.conditional_format(
                get_inside_tab_range(df_test, st_row=df_test_st_pos[0], st_col=df_test_st_pos[1]), {"type": "no_errors", "format": inner_format}
            )

            def set_aligment(df, st_row=st_row, st_col=st_col):
                col_pos = st_col
                if type(df.index[0]) == tuple:
                    col_pos = st_col + len(df.index[0])

                full_range = xu.xl_range(
                    st_row,
                    st_col,
                    st_row + len(df),
                    col_pos + len(df.columns) - 1,
                )
                worksheet.set_column(full_range, None, alignment)

            set_aligment(df)
            set_aligment(df_test, st_row=df_test_st_pos[0], st_col=df_test_st_pos[1])

            def get_indices_for_value(df, value):
                index_col_dlist = [[] for _ in range(len(df.index[0]))]
                for j in range(len(df.index)):
                    count = 0
                    for elem in df.index[j]:
                        if value == elem:
                            index_col_dlist[count].append(j + 1)
                        count += 1
                return index_col_dlist

            # H_row = 4
            red_format = workbook.add_format({"bg_color": "#ff9090"})
            green_format = workbook.add_format({"bg_color": "#90ff90"})
            yellow_format = workbook.add_format({"bg_color": "#ffff90"})
            gray_format = workbook.add_format({"bg_color": "#909090"})
            return

    # save_to_excel(df, st_col=2)
    # print(df)
    return

ExampleTwo()

