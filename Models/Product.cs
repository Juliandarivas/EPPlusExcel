using System;

namespace EPPlusExcel.Models
{
    public class Product
    {
        public int Id { get; }
        public string Referencia { get; }
        public string Descripcion { get; }
        public DateTime FechaCreacion { get; }
        public decimal Valor { get; }

        public Product(int id, string referencia, string descripcion, DateTime fechaCreacion, decimal valor)
        {
            Id = id;
            Referencia = referencia;
            Descripcion = descripcion;
            FechaCreacion = fechaCreacion;
            Valor = valor;
        }
    }
}