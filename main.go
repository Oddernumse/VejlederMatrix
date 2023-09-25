package main

import (
	"fmt"
	"math/rand"
)

type Vejleder struct {
	name  string
	tasks int
}

type Student struct {
	id         int
	teachers   [2]Vejleder
	groupScore int
}

var vejledere []Vejleder
var students []Student

func randInt(min, max int) int {
	return min + rand.Intn(max-min)
}

func createTeachers() []Vejleder {
	var vejlederNavne = []string{"børge", "george", "nikolaj", "john", "bent", "dorte", "anders", "sissel", "gundhilde", "jens"}

	vejledere = []Vejleder{}

	for i := 0; i < len(vejlederNavne); i++ {
		n := Vejleder{name: vejlederNavne[i], tasks: randInt(1, 10)}
		vejledere = append(vejledere, n)
	}

	return vejledere
}

func pairStudents(vejledere []Vejleder) []Student {

	students = []Student{}

	var elever int = 5

	//wb, err := xlsx.OpenFile("testing.xlsx")

	for i := 0; i < elever; i++ {
		var temp = [2]Vejleder{{name: vejledere[i+i].name, tasks: vejledere[i+i].tasks}, {name: vejledere[i+i+1].name, tasks: vejledere[i+i+1].tasks}}
		n := Student{teachers: temp, id: i, groupScore: vejledere[i+i].tasks + vejledere[i+i+1].tasks}
		students = append(students, n)
	}

	return students
}

func createBlocks(teachers []Vejleder, studentTeacher []Student) map[int][]Student {
	//varusedTeachers := []string{}

	var intervals = make(map[int][]Student)

	tempArray := []Student{}

	for i := 0; i < 12; i++ {
		intervals[i] = tempArray
	}

	return intervals
}

func main() {
	teachers := createTeachers()
	studentTeacher := pairStudents(teachers)
	//blocks := createBlocks(teachers, studentTeacher)

	fmt.Println(studentTeacher)
}
