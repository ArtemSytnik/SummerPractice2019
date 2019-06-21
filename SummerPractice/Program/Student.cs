using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Runtime.Serialization;

namespace Program
{
    [DataContract]
    public class Student
    {
        //ID
        [DataMember]
        public string id;
        //Фамилия студента
        [DataMember]
        public string last_name_ukr;
        //Имя студента
        [DataMember]
        public string name_ukr;
        //Номер группы
        [DataMember]
        public string group_number;
        //Название предмета
        [DataMember]
        public string short_name;
        //Балл
        [DataMember]
        public string name;
        //Форма проверки
        [DataMember]
        public string check_form;
        //Форма обучения
        [DataMember]
        public string name_1;
        //Фамилия преподователя
        [DataMember]
        public string last_name_ukr_1;
        //Имя преподователя
        [DataMember]
        public string name_ukr_1;
        [DataMember]
        public string chair_number;
        [DataMember]
        public string chair_number_1;
    }
}
