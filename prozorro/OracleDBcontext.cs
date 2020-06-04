using System;
using Microsoft.EntityFrameworkCore;

namespace prozorro
{
    public class OracleDBcontext:DbContext
    {
        public DbSet<Data> ProzorroParsedDatas { get; set; }
        private string ConnectionString { get; set; }
        public OracleDBcontext(string ConnectionString)
        {
            this.ConnectionString = ConnectionString;
            Database.EnsureCreated();
        }
        protected override void OnConfiguring(DbContextOptionsBuilder dbContextOptionsBuilder)
        {
            dbContextOptionsBuilder.UseOracle(ConnectionString);
        }
    }
}
