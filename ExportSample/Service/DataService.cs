using System.Collections.Generic;

namespace ExportSample.Service
{
    public static class DataService
    {
        #region ===[ Public Methods ]=============================================================

        /// <summary>
        /// Get Record List.
        /// </summary>
        /// <returns></returns>
        public static IEnumerable<DataProxy> GetData()
        {
            return new List<DataProxy>
                   {
                       new DataProxy
                           {
                               FirstName = "Sandeep",
                               LastName = "Kumar",
                               Email = "test1@email.com",
                               City = "Gurgaon"
                           },
                       new DataProxy
                           {
                               FirstName = "FName1",
                               LastName = "Lname1",
                               Email = "test2@email.com",
                               City = "New Delhi"
                           },
                       new DataProxy
                           {
                               FirstName = "FName4",
                               LastName = "Lname2",
                               Email = "test3@email.com",
                               City = "Mumbai"
                           },
                       new DataProxy
                           {
                               FirstName = "FName3",
                               LastName = "Lname3",
                               Email = "test4@email.com",
                               City = "Mumbai"
                           },
                       new DataProxy
                           {
                               FirstName = "FName4",
                               LastName = "Lnam4",
                               Email = "test5@email.com",
                               City = "Gurgaon"
                           }
                   };
        }

        #endregion
    }

    /// <summary>
    /// Proxy class
    /// </summary>
    public sealed class DataProxy
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public string City { get; set; }
    }
}