﻿// <auto-generated />
using System;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Metadata;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using wyplaty.Entities;

#nullable disable

namespace wyplaty.Migrations
{
    [DbContext(typeof(FrachtDbContext))]
    partial class FrachtDbContextModelSnapshot : ModelSnapshot
    {
        protected override void BuildModel(ModelBuilder modelBuilder)
        {
#pragma warning disable 612, 618
            modelBuilder
                .HasAnnotation("ProductVersion", "6.0.13")
                .HasAnnotation("Relational:MaxIdentifierLength", 128);

            SqlServerModelBuilderExtensions.UseIdentityColumns(modelBuilder, 1L, 1);

            modelBuilder.Entity("wyplaty.Entities.Driver", b =>
                {
                    b.Property<int>("Id")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int");

                    SqlServerPropertyBuilderExtensions.UseIdentityColumn(b.Property<int>("Id"), 1L, 1);

                    b.Property<string>("LastName")
                        .IsRequired()
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("Name")
                        .HasColumnType("nvarchar(max)");

                    b.HasKey("Id");

                    b.ToTable("Drivers");
                });

            modelBuilder.Entity("wyplaty.Entities.Fracht", b =>
                {
                    b.Property<int>("Id")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int");

                    SqlServerPropertyBuilderExtensions.UseIdentityColumn(b.Property<int>("Id"), 1L, 1);

                    b.Property<string>("Car")
                        .IsRequired()
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("CityLoad")
                        .IsRequired()
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("CityUnload")
                        .IsRequired()
                        .HasColumnType("nvarchar(max)");

                    b.Property<int?>("DriverId")
                        .HasColumnType("int");

                    b.Property<DateTime>("LoadDate")
                        .HasColumnType("datetime2");

                    b.Property<string>("OrderNumber")
                        .IsRequired()
                        .HasColumnType("nvarchar(450)");

                    b.Property<int>("Price")
                        .HasColumnType("int");

                    b.Property<DateTime>("UnloadDate")
                        .HasColumnType("datetime2");

                    b.HasKey("Id");

                    b.HasIndex("DriverId");

                    b.HasIndex("OrderNumber")
                        .IsUnique();

                    b.ToTable("Frachts");
                });

            modelBuilder.Entity("wyplaty.Entities.Fracht", b =>
                {
                    b.HasOne("wyplaty.Entities.Driver", "Driver")
                        .WithMany()
                        .HasForeignKey("DriverId");

                    b.Navigation("Driver");
                });
#pragma warning restore 612, 618
        }
    }
}
