using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace wyplaty.Migrations
{
    public partial class init : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Drivers",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Name = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    LastName = table.Column<string>(type: "nvarchar(max)", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Drivers", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "Frachts",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    OrderNumber = table.Column<string>(type: "nvarchar(450)", nullable: false),
                    LoadDate = table.Column<DateTime>(type: "datetime2", nullable: false),
                    UnloadDate = table.Column<DateTime>(type: "datetime2", nullable: false),
                    Car = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    CityLoad = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    CityUnload = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Price = table.Column<int>(type: "int", nullable: false),
                    DriverId = table.Column<int>(type: "int", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Frachts", x => x.Id);
                    table.ForeignKey(
                        name: "FK_Frachts_Drivers_DriverId",
                        column: x => x.DriverId,
                        principalTable: "Drivers",
                        principalColumn: "Id");
                });

            migrationBuilder.CreateIndex(
                name: "IX_Frachts_DriverId",
                table: "Frachts",
                column: "DriverId");

            migrationBuilder.CreateIndex(
                name: "IX_Frachts_OrderNumber",
                table: "Frachts",
                column: "OrderNumber",
                unique: true);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Frachts");

            migrationBuilder.DropTable(
                name: "Drivers");
        }
    }
}
