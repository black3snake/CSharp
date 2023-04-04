using System;
using System.Collections.Generic;
using System.Data.Entity.ModelConfiguration;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations.Schema;
using System.Runtime.Remoting.Contexts;
using System.Reflection.Emit;

namespace WForRestGet.DataModel
{
    public class DataModelContext : DbContext
    {
        /*static DataModelContext()
        {
        }*/
        public DataModelContext() : base("name=DBConnection")
        {
            Database.SetInitializer<DataModelContext>(new LeaveContextInitializer());
        }

        public DbSet<Datauser> Datausers { get; set; }
        public DbSet<Leave> Leaves { get; set; }

        #region OLD connection
        /*protected override void OnConfiguring(cont optionsBuilder)
        {
            optionsBuilder.UseSqlServer(@"Server=(localdb)\mssqllocaldb;Database=mobileappdb;Trusted_Connection=True;");
        }*/

        /*protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
                optionsBuilder.UseSqlServer("Data Source=SERVER; Initial Catalog=UsersLeave; TrustServerCertificate=True; Integrated Security=True;");
            }
        }*/
        #endregion 

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Configurations.Add(new DatauserConfigurations());
            modelBuilder.Configurations.Add(new LeaveConfigurations());

            #region OLD string ini Leave DB
            /*modelBuilder.Entity<Leave>();
                new Leave[]
                {
                    new Leave() { Id = 1, LeaveType = "VC", LeaveDescription = "отпуск" },
                    new Leave() { Id = 2, LeaveType= "SL", LeaveDescription= "больничный" },
                    new Leave() { Id = 3, LeaveType = "BT", LeaveDescription = "командировка"},
                    new Leave() { Id = 4, LeaveType = "DV", LeaveDescription = "декретный отпуск" }
                });*/
            #endregion
        }
        private class DatauserConfigurations : EntityTypeConfiguration<Datauser>   //IEntityTypeConfiguration<Datauser>
        {
            //public void Configure(EntityTypeBuilder<Datauser> builder)
            public DatauserConfigurations()
            {
                this.HasKey(e => e.FimSyncKey);
                this.Property(e => e.FimSyncKey).HasDatabaseGeneratedOption(DatabaseGeneratedOption.None); // .ValueGeneratedNever();
                this.Property(e => e.FimSyncKey).HasMaxLength(40).IsRequired();
                this.Property(e => e.AccountId).HasMaxLength(50).IsRequired();
                this.Property(e => e.AccountName).HasMaxLength(50);
                this.Property(e => e.LastName).HasMaxLength(100);
                this.Property(e => e.FirstName).HasMaxLength(100);
                this.Property(e => e.MiddleName).HasMaxLength(100);
                this.Property(e => e.EmployeeNumber).HasMaxLength(20);
                this.Property(e => e.Birthday).HasColumnType("date");
                this.Property(e => e.CompanyName).HasMaxLength(300);
                this.Property(e => e.DepartmentName).HasMaxLength(200);
                this.Property(e => e.JobTitle).HasMaxLength(200);
                this.Property(e => e.DateIn).HasColumnType("date");
                this.Property(e => e.LeaveId).HasDatabaseGeneratedOption(DatabaseGeneratedOption.Computed).HasColumnAnnotation("SqlDefaultValue", 0);   //.HasDefaultValue(0);
                this.Property(e => e.LeaveStart).HasColumnType("date");
                this.Property(e => e.LeaveEnd).HasColumnType("date");
                this.Property(e => e.City).HasMaxLength(100);
                this.Property(e => e.Phone).HasMaxLength(100);
                this.Property(e => e.Email).HasMaxLength(100);
                this.Property(e => e.Disabled).HasColumnType("bit");

                /*this.HasOne(l => l.Leave)
                    .WithMany(e => e.Datausers)
                .HasForeignKey(l => l.LeaveId);*/

            }
        }
        private class LeaveConfigurations : EntityTypeConfiguration<Leave>
        {
            //public void Configure(EntityTypeBuilder<Leave> builder)
            public LeaveConfigurations()
            {
                this.HasKey(e => e.Id);
                this.Property(e => e.Id).HasDatabaseGeneratedOption(DatabaseGeneratedOption.None);
                this.Property(e => e.LeaveType).HasMaxLength(4).IsRequired();
                this.Property(e => e.LeaveDescription).HasMaxLength(200).IsRequired();

                this.HasMany(e => e.Datausers)
                    .WithRequired(e => e.Leave);
            }
        }

    }
    class LeaveContextInitializer : DropCreateDatabaseIfModelChanges<DataModelContext>
    {
        protected override void Seed(DataModelContext context)
        {
            IList<Leave> leaves = new List<Leave>();

            leaves.Add(new Leave { Id = 1, LeaveType = "VC", LeaveDescription = "отпуск" });
            leaves.Add(new Leave { Id = 2, LeaveType = "SL", LeaveDescription = "больничный" });
            leaves.Add(new Leave { Id = 3, LeaveType = "BT", LeaveDescription = "командировка" });
            leaves.Add(new Leave { Id = 4, LeaveType = "DV", LeaveDescription = "декретный отпуск" });

            foreach (Leave leave in leaves)
                context.Leaves.Add(leave);
            base.Seed(context);
            context.SaveChanges();
        }
    }
}
