using Config.Net;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoOutlookRims.DataModel
{
    public class DataModelContext : DbContext
    {
        public DataModelContext()
        {
            Database.EnsureCreated();
        }
        public DataModelContext(DbContextOptions<DataModelContext> options)
            : base(options)
        {
        }

        public DbSet<Datauser> Datausers { get; set; }
        public DbSet<Leave> Leaves { get; set; }
        public DbSet<Answer> Answers { get; set; }

        public IMySettings configiniD = new ConfigurationBuilder<IMySettings>()
            .UseIniFile(@"config.ini", true)
            .Build();

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
                optionsBuilder.UseSqlServer($"Data Source={configiniD.Address}; Initial Catalog={configiniD.BDName}; TrustServerCertificate=True; Integrated Security=True;");
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.ApplyConfiguration(new DatauserConfigurations());
            modelBuilder.ApplyConfiguration(new LeaveConfigurations());
            modelBuilder.ApplyConfiguration(new AnswerConfigurations());

            modelBuilder.Entity<Leave>().HasData(
                new Leave[]
                {
                    new Leave() { Id = 1, LeaveType = "VC", LeaveDescription = "отпуск" },
                    new Leave() { Id = 2, LeaveType= "SL", LeaveDescription = "больничный" },
                    new Leave() { Id = 3, LeaveType = "BT", LeaveDescription = "командировка"},
                    new Leave() { Id = 4, LeaveType = "DV", LeaveDescription = "декретный отпуск" }
                });
            modelBuilder.Entity<Answer>().HasData(
                new Answer[]
                {
                    new Answer() { Id = 1, AnswerType = "AN", AnswerDescription = "Автоответа статус неизвестен" },
                    new Answer() { Id = 2, AnswerType = "AA", AnswerDescription = "Exception на установку ответа" },
                    new Answer() { Id = 3, AnswerType = "AG", AnswerDescription = "Автоответ установлен"},
                    new Answer() { Id = 4, AnswerType = "AU", AnswerDescription = "Автоответ установлен пользователем" }
                });

        }
        private class DatauserConfigurations : IEntityTypeConfiguration<Datauser>
        {
            public void Configure(EntityTypeBuilder<Datauser> builder)
            {
                builder.HasKey(e => e.FimSyncKey);
                builder.Property(e => e.FimSyncKey).ValueGeneratedNever();
                builder.Property(e => e.FimSyncKey).HasMaxLength(40).IsRequired();
                builder.Property(e => e.AccountId).HasMaxLength(50).IsRequired();
                builder.Property(e => e.AccountName).HasMaxLength(50);
                builder.Property(e => e.LastName).HasMaxLength(100);
                builder.Property(e => e.FirstName).HasMaxLength(100);
                builder.Property(e => e.MiddleName).HasMaxLength(100);
                builder.Property(e => e.EmployeeNumber).HasMaxLength(20);
                builder.Property(e => e.Birthday).HasColumnType("date");
                builder.Property(e => e.CompanyName).HasMaxLength(300);
                builder.Property(e => e.DepartmentName).HasMaxLength(200);
                builder.Property(e => e.JobTitle).HasMaxLength(200);
                builder.Property(e => e.DateIn).HasColumnType("date");
                builder.Property(e => e.LeaveId).HasDefaultValue(0);
                builder.Property(e => e.LeaveStart).HasColumnType("date");
                builder.Property(e => e.LeaveEnd).HasColumnType("date");
                builder.Property(e => e.City).HasMaxLength(100);
                builder.Property(e => e.Phone).HasMaxLength(100);
                builder.Property(e => e.Email).HasMaxLength(100);
                builder.Property(e => e.Disabled).HasColumnType("bit");
                builder.Property(e => e.AnswerId).HasDefaultValue(1);

                builder.HasOne(l => l.Leave)
                    .WithMany(e => e.Datausers)
                    .HasForeignKey(l => l.LeaveId);

                builder.HasOne(l => l.Answer)
                    .WithMany(e => e.Datausers)
                    .HasForeignKey(l => l.AnswerId);
            }
        }
        private class LeaveConfigurations : IEntityTypeConfiguration<Leave>
        {
            public void Configure(EntityTypeBuilder<Leave> builder)
            {
                builder.HasKey(e => e.Id);
                builder.Property(e => e.Id).ValueGeneratedNever();
                builder.Property(e => e.LeaveType).HasMaxLength(4).IsRequired();
                builder.Property(e => e.LeaveDescription).HasMaxLength(200).IsRequired();

            }
        }
        private class AnswerConfigurations : IEntityTypeConfiguration<Answer>
        {
            public void Configure(EntityTypeBuilder<Answer> builder)
            {
                builder.HasKey(e => e.Id);
                builder.Property(e => e.Id).ValueGeneratedNever();
                builder.Property(e => e.AnswerType).HasMaxLength(4).IsRequired();
                builder.Property(e => e.AnswerDescription).HasMaxLength(200).IsRequired();

            }
        }

    }

}
