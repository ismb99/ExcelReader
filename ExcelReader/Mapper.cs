namespace ExcelReader
{
    internal static class Mapper
    {

        public static List<Person> FromDtoMapper(List<PersonDto> personDtoList)
        {
            var mapToPersonList = new List<Person>();

            foreach (var personDto in personDtoList)
            {
                var person = new Person
                {
                    FirstName = personDto.FirstName,
                    LastName = personDto.LastName,
                    Age = personDto.Age,
                    City="N/A"
                };
                
                mapToPersonList.Add(person);
            }

            return mapToPersonList;
        }
    }
}