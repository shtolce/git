using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Patterns.Classes.FactoryMethod;

namespace Patterns
{
    class Program
    {
        static void Main(string[] args)
        {
            Product prA = new CreatorA().Create();
            Product prB = new CreatorB().Create();
            prA.Method1();
            prB.Method1();


        }



    }
}
