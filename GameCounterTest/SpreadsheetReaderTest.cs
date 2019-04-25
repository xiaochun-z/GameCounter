using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GameCounter;
using Xunit;

namespace GameCounterTest
{
    public class SpreadsheetReaderTest
    {
        [InlineData("Gears of War 4")]
        [InlineData("Zumba Fitness: World Party")]
        [InlineData("World of Tanks: Xbox One Edition")]
        [InlineData("Wolfenstein: The New Order")]
        [InlineData("#killallzombies")]
        [InlineData("2064: Read Only Memories")]
        [Theory]
        public void VerfiyXboxoneGameExists(string name)
        {
            var file = new FileInfo("library-04-25-2019.xlsx");
            var xboxoneGames = SpreadsheetReader.ReadGames(file, "Xboxone", 1, 2, 4, 7);

            Assert.Contains(xboxoneGames, game => game.Name == name);
            Assert.Equal(1952, xboxoneGames.Count);
        }

        [InlineData("0 Day Attack on Earth")]
        [InlineData("Zuma's Revenge!")]
        [Theory]
        public void VerfiyXbox360CompatibleExists(string name)
        {
            var file = new FileInfo("library-04-25-2019.xlsx");
            var xbox360CompatibleGames = SpreadsheetReader.ReadGames(file, "Xbox 360 Compatible", 1, -1, 2, 5);

            Assert.Contains(xbox360CompatibleGames, game => game.Name == name);
            Assert.Equal(555, xbox360CompatibleGames.Count);
        }

        [InlineData("Black")]
        [InlineData("Star Wars: Republic Commando")]
        [Theory]
        public void VerfiyXboxCompatibleExists(string name)
        {
            var file = new FileInfo("library-04-25-2019.xlsx");
            var xboxCompatibleGames = SpreadsheetReader.ReadGames(file, "Xbox 1st Gen Compatible", 1, -1, 2, 3);

            Assert.Contains(xboxCompatibleGames, game => game.Name == name);
            Assert.Equal(33, xboxCompatibleGames.Count);
        }

        [InlineData("#killallzombies")]
        [InlineData("Zotrix")]
        [InlineData("3on3 FreeStyle")]
        [InlineData("Pinball FX 3")]
        [InlineData("Lumo")]
        [InlineData("Machinarium")]
        [InlineData("God of War III Remastered")]
        [Theory]
        public void VerfiyPS4GameExists(string name)
        {
            var file = new FileInfo("library-04-25-2019.xlsx");
            var ps4Games = SpreadsheetReader.ReadGames(file, "PS4", 1, 2, 4, 7);

            Assert.Contains(ps4Games, game => game.Name == name);
            Assert.Equal(1964, ps4Games.Count);
        }
    }
}
