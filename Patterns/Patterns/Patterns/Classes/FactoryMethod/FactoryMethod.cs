using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Patterns.Classes.FactoryMethod
{
    /// <summary>
    /// This is exemplary product class for studying
    /// </summary>
    public abstract class Product
    {
        public Product(string name)
        {
            Console.WriteLine($"Конструктор {name}" ,name);
        }
        abstract public void Method1();
    }

    public class ConcreteProductA : Product
    {
        public string productName
        {
            get;set;
        }
        public ConcreteProductA() : base("A Product")
        {
        }
        public override void Method1()
        {
            Console.WriteLine("A Product");
        }

    }

    public class ConcreteProductB : Product
    {
        public string productName
        {
            get; set;
        }
        public ConcreteProductB() : base("B Product")
        {
        }
        public override void Method1()
        {
            Console.WriteLine("B Product");
        }

    }
    /// <summary>
    /// Examplary class of creating Product instances
    /// </summary>
    public abstract class Creator
    {
        abstract public Product Create();
    }
    public class CreatorA:Creator
    {
        public override Product Create()
        {
            return new ConcreteProductA();
        }

    }
    public class CreatorB : Creator
    {
        public override Product Create()
        {
            return new ConcreteProductB();
        }

    }





}
