using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReporteExcel02.Models
{
    public class DaoPeru
    {
        public List<departamentos> peruDepaProv()
        {
            List<departamentos> list = new List<departamentos>();

            using(var db=new ModelPeru())
            {
                var query = from depa in db.departamentos select depa;

                foreach(var depa in query)
                {
                    departamentos departamento = depa;

                    foreach(var prov in depa.provincias)
                    {
                        departamento.provincias.Add(prov);
                    }

                    list.Add(departamento);
                }
            }

            return list;
        }

        public List<departamentos> peruDepaProvDist()
        {
            List<departamentos> list = new List<departamentos>();

            using (var db = new ModelPeru())
            {
                var query = from depa in db.departamentos select depa;

                foreach (var depa in query)
                {
                    departamentos departamento = depa;

                    foreach (var prov in depa.provincias)
                    {
                        provincias provincia = prov;

                        foreach (var dist in prov.distritos)
                        {
                            provincia.distritos.Add(dist);
                        }

                        departamento.provincias.Add(prov);
                    }

                    list.Add(departamento);
                }
            }
            return list;
        }
    }
}