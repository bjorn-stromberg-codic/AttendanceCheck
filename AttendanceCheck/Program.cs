using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace AttendanceCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            string project_dir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

            // Fetch students
            var classes = new List<Class>();
            {
                var path = Path.Combine(project_dir, "StudentList.txt");
                var lines = File.ReadAllLines(path, Encoding.Default);

                Class active_class = null;
                foreach (var line in lines)
                {
                    if (line.Length == 0 || line[0] == '#')
                        continue;
                    else if (line[0] == '[')
                    {
                        active_class = new Class()
                        {
                            GroupName = line.Substring(1, line.IndexOf(']') - 1),
                            Names = new List<string>(),
                        };
                        classes.Add(active_class);
                    }
                    else
                    {
                        active_class.Names.Add(line);
                    }
                }
            }

            var csv_folder_path = Path.Combine(project_dir, "PutAllAttendanceCsvsHere");
            foreach (var dir in new DirectoryInfo(csv_folder_path).EnumerateDirectories())
            {
                var input_path = Path.Combine(csv_folder_path, dir.Name);
                var output_path = Path.Combine(project_dir, $"Närvarolista_WIN20_{dir.Name}.xlsx");
                PrintAttendanceSheet(input_path, output_path, classes);
            }
        }

        static void PrintAttendanceSheet(string input_path, string output_path, List<Class> classes) 
        {
            // Fetch attendance data
            var total_attendance = new List<Attendance>();
            {
                var dir_info = new DirectoryInfo(input_path);
                foreach(var file_info in dir_info.GetFiles("*.csv").OrderBy(f => f.LastWriteTime))
                {
                    var attendance = new Attendance()
                    {
                        Date = file_info.LastWriteTime,
                        Attendees = new HashSet<string>(),
                    };

                    var lines = File.ReadAllLines(file_info.FullName);
                    for (int i = 1; i < lines.Length; i++)
                    {
                        var first_cell = lines[i].Split('\t')[0];
                        attendance.Attendees.Add(first_cell);
                    }

                    total_attendance.Add(attendance);
                }
            }

            // Print excel result sheet
            {
                // Boot Excel
                var instance = new Excel.Application();
                instance.Visible = false;
                instance.DisplayAlerts = false;
                var workbook = instance.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.Worksheets.Add();
                
                // Set base background color
                Excel.Range big_chunk = worksheet.Range[
                                    worksheet.Cells[1, 1],
                                    worksheet.Cells[100, 100]
                                ];
                big_chunk.Interior.Color = Excel.XlRgbColor.rgbOrange;
                big_chunk.Interior.TintAndShade = 0.8;

                // Fill cells with data
                {
                    var top_row_offset = 1;
                    var attendance_col_offset = 4;

                    for (int i = 0; i < total_attendance.Count; i++)
                    {
                        var attendance = total_attendance[i];
                        var cell = worksheet.Cells[top_row_offset, attendance_col_offset + i];
                        cell.Value = "'" + attendance.Date.ToString("dd/MM", CultureInfo.InvariantCulture);
                        cell.Interior.TintAndShade = 0.6;
                    }

                    var start_row_offset = top_row_offset;
                    foreach (var @class in classes)
                    {
                        PrintStudentAttendance(worksheet, @class, start_row_offset, total_attendance, attendance_col_offset);
                        start_row_offset += @class.Names.Count + 2;
                    }
                }

                // Save Excel sheet
                workbook.SaveAs(output_path);
                workbook.Close();
                instance.Quit();
            }
        }
        static void PrintStudentAttendance(Excel.Worksheet worksheet, Class @class, int row_offset, List<Attendance> total_attendance, int attendance_col_offset)
        {
            worksheet.Cells[row_offset, 1].Value = @class.GroupName;
            worksheet.Cells[row_offset, 1].Interior.TintAndShade = 0.6;

            bool[] somebody_attended = new bool[total_attendance.Count];
            for (int i = 0; i < @class.Names.Count; i++)
            {
                var row = row_offset + 1 + i;
                var name = @class.Names[i];
                worksheet.Cells[row, 1].Value = name;

                var attended_days = 0;
                for (int j = 0; j < total_attendance.Count; j++)
                {
                    var attendance = total_attendance[j];
                    if(attendance.Attendees.Contains(name))
                    {
                        attended_days++;
                        worksheet.Cells[row, attendance_col_offset + j] = "X";
                        worksheet.Cells[row, attendance_col_offset + j].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        somebody_attended[j] = true;
                    }
                }
                worksheet.Cells[row, attendance_col_offset - 1].Value = attended_days;
            }

            for (int j = 0; j < total_attendance.Count; j++)
            {
                if (!somebody_attended[j])
                {
                    Excel.Range big_chunk = worksheet.Range[
                                                worksheet.Cells[row_offset + 1,                      attendance_col_offset + j],
                                                worksheet.Cells[row_offset + 1 + @class.Names.Count, attendance_col_offset + j]
                                            ];
                    big_chunk.Interior.Color = Excel.XlRgbColor.rgbAzure;
                }
            }
        }
        class Attendance
        {
            public DateTime Date;
            public HashSet<string> Attendees;
        }

        class Class
        {
            public List<string> Names;
            public string GroupName;
        }
    }
}
