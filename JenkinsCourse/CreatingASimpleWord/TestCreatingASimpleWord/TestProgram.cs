using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using CreatingASimpleWord;

namespace TestCreatingASimpleWord
{
    [TestFixture]
    class TestProgram
    {
       [Test]
        public void ValidateMessage()
       {
           Assert.AreEqual("Hello World!", Program.CreateMessage(), "Mesaj incorect");
       }
    }
}
