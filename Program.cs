using System;

namespace TeamsAutoJoiner
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Teams t = new Teams();
                t.Work(args);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}
