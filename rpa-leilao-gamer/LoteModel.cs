using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace rpa_leilao_gamer
{
    public class LoteModel
    {
        public string Link { get; set; }
        public string Lote { get; set; }
        public string Description { get; set; }
        public string MaxCast { get; set; }

        public LoteModel(string lote, string description,
            string maxCast, string link)
        {
            this.Lote = lote;
            this.Description = description;
            this.MaxCast = maxCast;  
            this.Link = link;
        }
    }
}
