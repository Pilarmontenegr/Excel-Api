using Excel_Api.Controllers;
using OfficeOpenXml;

namespace Excel_Api.Rules
{
    public class Sueldo
    {
        public int MontoAnual { get; set; }
        public int Monto { get; set; }
        public string Nombre { get; set; }

        public byte[] CrearExcel()
        {

            using (var package = new ExcelPackage())
            {
                ExcelWorkbook excelWorkbook = package.Workbook;
                
                
                    var sheet = package.Workbook.Worksheets.Add("Sheet 1");



                    sheet.Cells["A1"].Value = "Nombre";
                    var nombreIngresado = Nombre;
                    sheet.Cells["A2"].Value = nombreIngresado;


                    sheet.Cells["B1"].Value = "Sueldo Mensual";
                    //sheet.Cells["B1"].Width = 15;
                    var montoIngresado = Monto;
                    var montoIngresadoAnual = MontoAnual / 12;
                    if (montoIngresado > 0)
                        sheet.Cells["B2"].Value = montoIngresado;
                    else sheet.Cells["B2"].Value = montoIngresadoAnual;


                    sheet.Cells["C1"].Value = "Enero";
                    if (montoIngresado > 0)
                        sheet.Cells["C2"].Value = montoIngresado;
                    else sheet.Cells["C2"].Value = montoIngresadoAnual;

                    sheet.Cells["D1"].Value = "Febrero";
                    if (montoIngresado > 0)
                        sheet.Cells["D2"].Value = montoIngresado;
                    else sheet.Cells["D2"].Value = montoIngresadoAnual;


                    sheet.Cells["E1"].Value = "Marzo";
                    if (montoIngresado > 0)
                        sheet.Cells["E2"].Value = montoIngresado;
                    else sheet.Cells["E2"].Value = montoIngresadoAnual;

                    sheet.Cells["F1"].Value = "Abril";
                    if (montoIngresado > 0)
                        sheet.Cells["F2"].Value = montoIngresado;
                    else sheet.Cells["F2"].Value = montoIngresadoAnual;

                    sheet.Cells["G1"].Value = "Mayo";
                    if (montoIngresado > 0)
                        sheet.Cells["G2"].Value = montoIngresado;
                    else sheet.Cells["G2"].Value = montoIngresadoAnual;


                    sheet.Cells["H1"].Value = "Junio";
                    if (montoIngresado > 0)
                        sheet.Cells["H2"].Value = montoIngresado;
                    else sheet.Cells["H2"].Value = montoIngresadoAnual;


                    sheet.Cells["I1"].Value = "S.A.C";
                    if (montoIngresado > 0)
                        sheet.Cells["I2"].Value = montoIngresado / 2;
                    else sheet.Cells["I2"].Value = montoIngresadoAnual / 2;


                    sheet.Cells["J1"].Value = "Julio";
                    if (montoIngresado > 0)
                        sheet.Cells["J2"].Value = montoIngresado;
                    else sheet.Cells["J2"].Value = montoIngresadoAnual;


                    sheet.Cells["K1"].Value = "Agosto";
                    if (montoIngresado > 0)
                        sheet.Cells["K2"].Value = montoIngresado;
                    else sheet.Cells["K2"].Value = montoIngresadoAnual;

                    sheet.Cells["L1"].Value = "Septiembre";
                    if (montoIngresado > 0)
                        sheet.Cells["L2"].Value = montoIngresado;
                    else sheet.Cells["L2"].Value = montoIngresadoAnual;


                    sheet.Cells["M1"].Value = "Octubre";
                    if (montoIngresado > 0)
                        sheet.Cells["M2"].Value = montoIngresado;
                    else sheet.Cells["M2"].Value = montoIngresadoAnual;

                    sheet.Cells["N1"].Value = "Noviembre";
                    if (montoIngresado > 0)
                        sheet.Cells["N2"].Value = montoIngresado;
                    else sheet.Cells["N2"].Value = montoIngresadoAnual;

                    sheet.Cells["O1"].Value = "Diciembre";
                    if (montoIngresado > 0)
                        sheet.Cells["O2"].Value = montoIngresado;
                    else sheet.Cells["O2"].Value = montoIngresadoAnual;

                    sheet.Cells["P1"].Value = "S.A.C";
                    if (montoIngresado > 0)
                        sheet.Cells["P2"].Value = montoIngresado / 2;
                    else sheet.Cells["P2"].Value = montoIngresadoAnual / 2;

                    sheet.Cells["Q1"].Value = "Descuento Ganancias";
                    if (Monto > 300000 || MontoAnual > 3600000)
                    {
                        sheet.Cells["Q2"].Value = "SI";
                    }
                        
                    else
                    sheet.Cells["Q2"].Value = "NO";

                    return package.GetAsByteArray();

                    //package.Save();
                

            }
                

        }
    }
}
//using (var package = new ExcelPackage(@"C:\Users\Pili\Documents\piliColegio\Calero\formExcel.xlsx"))
//{

//    var sheet = package.Workbook.Worksheets.Add("Sheet 1");



//    sheet.Cells["A1"].Value = "Nombre";
//    var nombreIngresado = Nombre;
//    sheet.Cells["A2"].Value = nombreIngresado;


//    sheet.Cells["B1"].Value = "Sueldo Mensual";
//    //sheet.Cells["B1"].Width = 15;
//    var montoIngresado = Monto;
//    var montoIngresadoAnual = MontoAnual / 12;
//    if (montoIngresado > 0)
//        sheet.Cells["B2"].Value = montoIngresado;
//    else sheet.Cells["B2"].Value = montoIngresadoAnual;


//    sheet.Cells["C1"].Value = "Enero";
//    if (montoIngresado > 0)
//        sheet.Cells["C2"].Value = montoIngresado;
//    else sheet.Cells["C2"].Value = montoIngresadoAnual;

//    sheet.Cells["D1"].Value = "Febrero";
//    if (montoIngresado > 0)
//        sheet.Cells["D2"].Value = montoIngresado;
//    else sheet.Cells["D2"].Value = montoIngresadoAnual;


//    sheet.Cells["E1"].Value = "Marzo";
//    if (montoIngresado > 0)
//        sheet.Cells["E2"].Value = montoIngresado;
//    else sheet.Cells["E2"].Value = montoIngresadoAnual;

//    sheet.Cells["F1"].Value = "Abril";
//    if (montoIngresado > 0)
//        sheet.Cells["F2"].Value = montoIngresado;
//    else sheet.Cells["F2"].Value = montoIngresadoAnual;

//    sheet.Cells["G1"].Value = "Mayo";
//    if (montoIngresado > 0)
//        sheet.Cells["G2"].Value = montoIngresado;
//    else sheet.Cells["G2"].Value = montoIngresadoAnual;


//    sheet.Cells["H1"].Value = "Junio";
//    if (montoIngresado > 0)
//        sheet.Cells["H2"].Value = montoIngresado;
//    else sheet.Cells["H2"].Value = montoIngresadoAnual;


//    sheet.Cells["I1"].Value = "S.A.C";
//    if (montoIngresado > 0)
//        sheet.Cells["I2"].Value = montoIngresado / 2;
//    else sheet.Cells["I2"].Value = montoIngresadoAnual / 2;


//    sheet.Cells["J1"].Value = "Julio";
//    if (montoIngresado > 0)
//        sheet.Cells["J2"].Value = montoIngresado;
//    else sheet.Cells["J2"].Value = montoIngresadoAnual;


//    sheet.Cells["K1"].Value = "Agosto";
//    if (montoIngresado > 0)
//        sheet.Cells["K2"].Value = montoIngresado;
//    else sheet.Cells["K2"].Value = montoIngresadoAnual;

//    sheet.Cells["L1"].Value = "Septiembre";
//    if (montoIngresado > 0)
//        sheet.Cells["L2"].Value = montoIngresado;
//    else sheet.Cells["L2"].Value = montoIngresadoAnual;


//    sheet.Cells["M1"].Value = "Octubre";
//    if (montoIngresado > 0)
//        sheet.Cells["M2"].Value = montoIngresado;
//    else sheet.Cells["M2"].Value = montoIngresadoAnual;

//    sheet.Cells["N1"].Value = "Noviembre";
//    if (montoIngresado > 0)
//        sheet.Cells["N2"].Value = montoIngresado;
//    else sheet.Cells["N2"].Value = montoIngresadoAnual;

//    sheet.Cells["O1"].Value = "Diciembre";
//    if (montoIngresado > 0)
//        sheet.Cells["O2"].Value = montoIngresado;
//    else sheet.Cells["O2"].Value = montoIngresadoAnual;

//    sheet.Cells["P1"].Value = "S.A.C";
//    if (montoIngresado > 0)
//        sheet.Cells["P2"].Value = montoIngresado / 2;
//    else sheet.Cells["P2"].Value = montoIngresadoAnual / 2;

//    sheet.Cells["Q1"].Value = "Descuento Ganancias";
//    if (Monto > 300000 || MontoAnual > 3600000)
//        sheet.Cells["Q2"].Value = "SI";

//    package.Save();
//}

//        }