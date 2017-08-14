using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MG_BoletaDefense
{
    public class Boleta : InterfaceBoleta
    {
        public void GenerarBoleta(List<IBoleta> ListaCabecera, List<TBoleta> ListaDetalle,string Ruta)
        {
            foreach (var cabecera in ListaCabecera)
            {
                var Detalles = ListaDetalle.Where(x => x.CodigoTrabajador == cabecera.CodigoTrabajador).ToList();

                using (Document oDocument = new Document())
                {
                    oDocument.SetPageSize(PageSize.A4.Rotate());
                    using (FileStream oFileStream = new FileStream(string.Format(@"{0}\{1}.pdf", Ruta, cabecera.CodigoTrabajador), FileMode.OpenOrCreate))
                    {
                        //Definimos la fuente y el tamaño
                        // BaseFont bfTimes = BaseFont.CreateFont(BaseFont., BaseFont.CP1252, false);
                        FontFactory.RegisterDirectories();
                        Font fontArialBold = new Font(FontFactory.GetFont("Arial", 7, Font.BOLD));
                        Font fontArialNormal = new Font(FontFactory.GetFont("Arial", 7, Font.NORMAL));
                        // Font times = new Font(bfTimes, 7);
                        //-->

                        PdfWriter.GetInstance(oDocument, oFileStream);
                        oDocument.Open();

                        string rutaImagen = @"E:\BoletasPdf\Logo.png";
                        // Creamos la imagen y le ajustamos el tamaño
                        Image imagen =Image.GetInstance(rutaImagen);
                        imagen.BorderWidth = 0;
                        imagen.Alignment = Element.ALIGN_LEFT;
                        float percentage = 0.0f;
                        percentage = 150 / imagen.Width;
                        imagen.ScalePercent(percentage * 100);
                        oDocument.Add(imagen);

                        PdfPTable Boleta = new PdfPTable(2);
                        //Medidas de la tabla
                        float[] medidaCeldas = { 303f, 500f };
                        Boleta.TotalWidth = 803f;
                        Boleta.LockedWidth = true;
                        Boleta.SetWidths(medidaCeldas);

                        //8 columnas para la Primera Fila
                        PdfPTable PrimerFila = new PdfPTable(8);
                        //Medidas de columnas primera fila
                        float[] medidaPrimeraFila = { 40f, 40f, 30f, 247f, 70f, 100f, 60f, 216f };
                        PrimerFila.TotalWidth = 803f;
                        PrimerFila.LockedWidth = true;
                        PrimerFila.SetWidths(medidaPrimeraFila);

                        //
                        PdfPCell cell;
                        //
                        cell = new PdfPCell(new Phrase("C. COSTO", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        PrimerFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("CÓDIGO", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        PrimerFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("TIPO", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        PrimerFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("APELLIDOS Y NOMBRES", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        PrimerFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("FECHA INGRESO", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        PrimerFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("CÓDIGO AFP", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        PrimerFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("L.E. / DNI", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        PrimerFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("DOMICILIO", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        PrimerFila.AddCell(cell);

                        //->Detalles Primera Fila
                        cell = new PdfPCell(new Phrase(String.Format(cabecera.CodCentroCosto), fontArialNormal));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        PrimerFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(String.Format(cabecera.CodigoTrabajador), fontArialNormal));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        PrimerFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(String.Format(cabecera.TipoPlanilla), fontArialNormal));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        PrimerFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(String.Format(cabecera.DesTrabajador), fontArialNormal));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        PrimerFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(String.Format(cabecera.FeIngreso.ToShortDateString()), fontArialNormal));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        PrimerFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(String.Format(cabecera.Afp), fontArialNormal));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        PrimerFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(String.Format(cabecera.NumeroDoc), fontArialNormal));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        PrimerFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(String.Format(cabecera.Direccion), fontArialNormal));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        PrimerFila.AddCell(cell);
                        //->

                        //Segunda Fila
                        cell = new PdfPCell(PrimerFila);
                        cell.Colspan = 2;
                        cell.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                        Boleta.AddCell(cell);


                        //10 columnas para la Primera Fila
                        PdfPTable SegundaFila = new PdfPTable(10);
                        //Medidas de columnas segunda fila
                        float[] medidaSegundaFila = { 150f, 50f, 50f, 100f, 100f, 100f, 95f, 30f, 75f, 53f };
                        SegundaFila.TotalWidth = 803f;
                        SegundaFila.LockedWidth = true;
                        SegundaFila.SetWidths(medidaSegundaFila);

                        //---
                        cell = new PdfPCell(new Phrase("CARGO", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("D. TRAB", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("H. TRAB", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("REMUNER. BASICA", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("VAC. SALIDA", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("VAC. RETORNO", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("CARNET ESSALUD", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("EDAD", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("FECHA CESE", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("CONDICIÓN", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        //->Agregamos detalles de la segunda fila

                        cell = new PdfPCell(new Phrase(String.Format(cabecera.CargoTrabajador), fontArialNormal));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(String.Format(cabecera.DiasTrabajados.ToString("###,##0.00")), fontArialNormal));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(String.Format(cabecera.HorasTrabajadas.ToString("###,##0.00")), fontArialNormal));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(String.Format(cabecera.SueldoBasico.ToString("###,##0.00")), fontArialNormal));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(String.Format(cabecera.FeSalidaVac.ToShortDateString()), fontArialNormal));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(String.Format(cabecera.FeIngresoVac.ToShortDateString()), fontArialNormal));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(String.Format(cabecera.Ips), fontArialNormal));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(String.Format(cabecera.Edad.ToString()), fontArialNormal));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(String.Format(cabecera.FeCese.ToShortDateString()), fontArialNormal));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(String.Format(cabecera.Situacion), fontArialNormal));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Border = Rectangle.RIGHT_BORDER;
                        SegundaFila.AddCell(cell);

                        //->


                        cell = new PdfPCell(SegundaFila);
                        cell.Colspan = 2;
                        cell.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                        Boleta.AddCell(cell);
                        //------------------Tercera Fila
                        PdfPTable TerceraFila = new PdfPTable(2);
                        float[] medidaTerceraFila = { 303f, 500f };
                        TerceraFila.TotalWidth = 803f;
                        TerceraFila.LockedWidth = true;
                        TerceraFila.SetWidths(medidaCeldas);

                        cell = new PdfPCell(new Phrase("REMUNERACIONES", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        TerceraFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("RETENCIONES", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER; //0=Left, 1=Centre, 2=Right
                        TerceraFila.AddCell(cell);

                        cell = new PdfPCell(TerceraFila);
                        cell.Colspan = 2;
                        cell.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                        Boleta.AddCell(cell);
                        //-----------------------------------------------
                        //Cuarta Fila
                        PdfPTable CuartaFila = new PdfPTable(8);

                        float[] medidaCuartaFila = { 163f, 70f, 70f, 160f, 80f, 80f, 80f, 100f };
                        CuartaFila.TotalWidth = 803f;
                        CuartaFila.LockedWidth = true;
                        CuartaFila.SetWidths(medidaCuartaFila);

                        //Aqui se insertan los detalles
                        cell = new PdfPCell(new Phrase("CONCEPTO", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cell.Rowspan = 2;
                        CuartaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("TIEMPO", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cell.Rowspan = 2;
                        CuartaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("MONTO", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.VerticalAlignment = Element.ALIGN_MIDDLE; ;
                        cell.Rowspan = 2;
                        CuartaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("CONCEPTO", fontArialBold));
                        cell.Rowspan = 2;
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.VerticalAlignment = Element.ALIGN_MIDDLE; ;
                        CuartaFila.AddCell(cell);

                        //CUARTA SUB TABLA   
                        cell = new PdfPCell(new Phrase("CUENTAS CORRIENTES", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Colspan = 3;
                        CuartaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("DESCUENTO", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        CuartaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("CARGO", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        CuartaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("AMORTIZACIÓN", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        CuartaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("SALDO", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        CuartaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("MONTO", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        CuartaFila.AddCell(cell);


                        //->Detalles del cuerpo de la boleta
                        var DetallesRemuneraciones = Detalles.Where(x => x.TipoConcepto == "1").ToList();
                        var DetallesRetenciones = Detalles.Where(x => x.TipoConcepto == "2").ToList();

                        int count = Detalles.Count;

                        for (int i  = 0; i < Detalles.Count; i++ )
                        {
                            if(i < DetallesRemuneraciones.Count)
                            {
                                cell = new PdfPCell(new Phrase(DetallesRemuneraciones[i].DesConcepto, fontArialNormal));
                                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                cell.Border = Rectangle.RIGHT_BORDER;
                                CuartaFila.AddCell(cell);

                                cell = new PdfPCell(new Phrase(string.Format(DetallesRemuneraciones[i].DiasHoras.ToString("###,##0.00")) , fontArialNormal));
                                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                cell.Border = Rectangle.RIGHT_BORDER;
                                CuartaFila.AddCell(cell);

                                cell = new PdfPCell(new Phrase(string.Format(DetallesRemuneraciones[i].ImporteConcepto.ToString("###,##0.00")), fontArialNormal));
                                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                cell.Border = Rectangle.RIGHT_BORDER;
                                CuartaFila.AddCell(cell);
                            }
                            else
                            {
                                for (int j = 0; j < 3; j++)
                                {
                                    cell = new PdfPCell(new Phrase(" ", fontArialNormal));
                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                    cell.Border = Rectangle.RIGHT_BORDER;
                                    CuartaFila.AddCell(cell);
                                }
                            }
                            
                            if(i < DetallesRetenciones.Count)
                            {
                                cell = new PdfPCell(new Phrase(DetallesRetenciones[i].DesConcepto, fontArialNormal));
                                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                cell.Border = Rectangle.RIGHT_BORDER;
                                CuartaFila.AddCell(cell);

                                cell = new PdfPCell(new Phrase(string.Format(DetallesRetenciones[i].Cargo.ToString("###,##0.00")), fontArialNormal));
                                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                cell.Border = Rectangle.RIGHT_BORDER;
                                CuartaFila.AddCell(cell);

                                cell = new PdfPCell(new Phrase(string.Format(DetallesRetenciones[i].Amortizacion.ToString("###,##0.00")), fontArialNormal));
                                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                cell.Border = Rectangle.RIGHT_BORDER;
                                CuartaFila.AddCell(cell);

                                cell = new PdfPCell(new Phrase(string.Format(DetallesRetenciones[i].Saldo.ToString("###,##0.00")), fontArialNormal));
                                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                cell.Border = Rectangle.RIGHT_BORDER;
                                CuartaFila.AddCell(cell);

                                cell = new PdfPCell(new Phrase(string.Format(DetallesRetenciones[i].ImporteConcepto.ToString("###,##0.00")), fontArialNormal));
                                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                cell.Border = Rectangle.RIGHT_BORDER;
                                CuartaFila.AddCell(cell);
                            }
                            else
                            {
                                for (int j = 0; j < 5; j++)
                                {
                                    cell = new PdfPCell(new Phrase(" ", fontArialNormal));
                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                    cell.Border = Rectangle.RIGHT_BORDER;
                                    CuartaFila.AddCell(cell);
                                }
                            }
                        }

                        CuartaFila.Rows.Count();
                        int filasRestantes = 20 - CuartaFila.Rows.Count();
                        for (int i = 0; i < filasRestantes; i++)
                        {
                            for (int j = 0; j < 8; j++)
                            {
                                cell = new PdfPCell(new Phrase(" ", fontArialNormal));
                                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                cell.Border = Rectangle.RIGHT_BORDER;
                                CuartaFila.AddCell(cell);
                            }   
                        }
                        //->Fin de los detalles de cuerpo

                        cell = new PdfPCell(CuartaFila);
                        cell.Colspan = 2;
                        cell.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                        Boleta.AddCell(cell);
                        //----------


                        //-----Quinta Fila
                        PdfPTable QuintaFila = new PdfPTable(3);
                        float[] medidaQuintaFila = { 303f, 390f, 110f };
                        QuintaFila.TotalWidth = 803f;
                        QuintaFila.LockedWidth = true;
                        QuintaFila.SetWidths(medidaQuintaFila);

                        cell = new PdfPCell(new Phrase("TOTAL REMUNERACIONES", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        QuintaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("TOTAL RETENCIONES", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        QuintaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(""));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        QuintaFila.AddCell(cell);

                        cell = new PdfPCell(QuintaFila);
                        cell.Colspan = 2;
                        cell.HorizontalAlignment = 1;

                        Boleta.AddCell(cell);
                        //-------

                        //------Sexta Fila
                        PdfPTable SextaFila = new PdfPTable(7);
                        float[] medidaSextaFila = { 75.75f, 75.75f, 75.75f, 75.75f, 166.6666666666667f, 166.6666666666667f, 166.6666666666667f };
                        SextaFila.TotalWidth = 803f;
                        SextaFila.LockedWidth = true;
                        SextaFila.SetWidths(medidaSextaFila);

                        cell = new PdfPCell(new Phrase("CONTRIBUCIONES PATRONALES", fontArialBold));
                        cell.Colspan = 4;
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        SextaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(""));
                        cell.Rowspan = 2;
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        SextaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("RECIBI CONFORME", fontArialBold));
                        cell.Rowspan = 2;
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        SextaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("S/.", fontArialBold));
                        cell.Rowspan = 3;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        cell.VerticalAlignment = Element.ALIGN_BOTTOM;
                        SextaFila.AddCell(cell);

                        //---
                        cell = new PdfPCell(new Phrase("ESSALUD", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        SextaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("ONP", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        SextaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(""));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        SextaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("IMP SOLES", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        SextaFila.AddCell(cell);

                        //----

                        cell = new PdfPCell(new Phrase(""));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        SextaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(""));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        SextaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(""));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        SextaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase(""));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        SextaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("FIRMA DEL REPRESENTANTE", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        SextaFila.AddCell(cell);

                        cell = new PdfPCell(new Phrase("FIRMA DEL TRABAJADOR", fontArialBold));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        SextaFila.AddCell(cell);

                        //----
                        cell = new PdfPCell(SextaFila);
                        cell.Colspan = 2;
                        cell.HorizontalAlignment = 1;
                        Boleta.AddCell(cell);
                        //-------

                        oDocument.Add(Boleta);

                        oDocument.Close();
                    }
                }   
            }
        }
    }
}
