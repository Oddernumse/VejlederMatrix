package main

import (
	"fmt"
	"math/rand"
	"sort"
)

type Vejleder struct {
	name string
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

func remove[T comparable](slice []T, s int) []T {
	return append(slice[:s], slice[s+1:]...)
}

func findAndDelete(s [4]int, item int) []int {
	index := 0
	for _, i := range s {
		if i != item {
			s[index] = i
			index++
		}
	}
	return s[:index]
}

func doesElementExist(s []string, str string) bool {
	for _, v := range s {
		if v == str {
			return true
		}
	}
	return false
}

func createTeachers() []Vejleder {
	var vejlederNavne = []string{"børge", "george", "nikolaj", "john", "børge", "dorte", "anders", "sissel", "gundhilde", "nikolaj"}

	vejledere = []Vejleder{}

	for i := 0; i < len(vejlederNavne); i++ {
		n := Vejleder{name: vejlederNavne[i]}
		vejledere = append(vejledere, n)
	}

	return vejledere
}

func pairStudents(vejledere []Vejleder) []Student {

	students = []Student{}

	var elever int = 5

	//wb, err := xlsx.OpenFile("testing.xlsx")

	for i := 0; i < elever; i++ {
		var temp = [2]Vejleder{{name: vejledere[i+i].name}, {name: vejledere[i+i+1].name}}
		n := Student{teachers: temp, id: i, groupScore: randInt(0, 10)}
		students = append(students, n)
	}

	return students
}

func createBlocks(teachers []Vejleder, studentTeacher []Student) map[int][]Student {
	var intervals = make(map[int][]Student)

	for i := 0; i < 12; i++ {

		studentTeacher := studentTeacher

		// We have some pretty big issues here
		// It always takes the first in the list and decreases its groupScore by one but thats not enough to make the next group more
		// Prioritised so it always takes the same group every time

		// Fix pls

		tempArray := []Student{}

		possible := true
		usedTeachers := []string{}
		for possible {

			fmt.Println(studentTeacher)
			if len(studentTeacher) > 0 && !doesElementExist(usedTeachers, studentTeacher[0].teachers[0].name) && !doesElementExist(usedTeachers, studentTeacher[0].teachers[1].name) {
				tempArray = append(tempArray, studentTeacher[0])
				usedTeachers = append(usedTeachers, studentTeacher[0].teachers[0].name)
				usedTeachers = append(usedTeachers, studentTeacher[0].teachers[1].name)

				//fmt.Println(usedTeachers, i)

				if studentTeacher[0].groupScore != 0 {
					studentTeacher[0].groupScore -= 1
				} else if studentTeacher[0].groupScore == 0 {
					studentTeacher = studentTeacher[1:]
				}

				intervals[i] = tempArray

				sort.Slice(studentTeacher, func(i, j int) bool {
					return studentTeacher[i].groupScore > studentTeacher[j].groupScore
				})
			} else {
				possible = false
			}

		}

	}

	return intervals

}

func main() {
	teachers := createTeachers()
	studentTeacher := pairStudents(teachers)
	blocks := createBlocks(teachers, studentTeacher)

	for k, v := range blocks {
		fmt.Printf("Value: %s\n", k, v)
	}
	fmt.Println(studentTeacher)
}
