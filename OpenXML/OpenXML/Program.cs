﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace OpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Report report = new Report();

                var inscriptos = new List<InscritoDTO>();
                Random random = new Random();
                for (int i = 0; i < 10; i++)
                {
                    inscriptos.Add(new InscritoDTO()
                    {
                        Id=i,
                        Num_socio = i,
                        Dni = 1000000 + i,
                        Nombre = "Inscripto " + i,
                        Email = "mail " + i,
                        Telefono = 1000000 + i,
                        Fecha_reg = "fecha ",
                        Lugar = "lugar " + i,
                        Horario = "horario " + i
                    });
                }

                byte [] excel= report.CreateExcelDoc(inscriptos);
                
                File.WriteAllBytes(@"d:\Reporte.xlsx", excel);                 
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }
    }
}
