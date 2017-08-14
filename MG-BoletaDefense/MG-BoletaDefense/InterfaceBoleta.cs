using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MG_BoletaDefense
{
    public interface InterfaceBoleta
    {
        void GenerarBoleta(List<IBoleta> ListaCabecera, List<TBoleta> ListaDetalle, string Ruta);
    }
}
