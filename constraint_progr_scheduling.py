from ortools.sat.python import cp_model
from pandas import DataFrame, Series

name_rooms = ["Kitchen1", "Kitchen2", "Living Room", "Hallway", "Bathroom1", "Bathroom2"]
name_persons = ["Name1", "Name2", "Name3", "Name4", "Name5", "Name6"]


class VarArraySolutionPrinterWithLimit(cp_model.CpSolverSolutionCallback):

    def __init__(self, limit):
        cp_model.CpSolverSolutionCallback.__init__(self)
        self.__solution_count = 0
        self.__solution_limit = limit

    def on_solution_callback(self):
        self.__solution_count += 1
        if self.__solution_count >= self.__solution_limit:
            print(f"Stop search after {self.__solution_limit} solutions")
            self.StopSearch()

    def solution_count(self):
        return self.__solution_count


def main():
    num_rooms = 6
    num_people = 6
    num_weeks = 16 # because it repeats itself after 16 weeks

    all_rooms = range(num_rooms)
    all_people = range(num_people)
    all_weeks = range(num_weeks)

    model = cp_model.CpModel()

    # Variables
    assignments = {}
    for week in all_weeks:
        for room in all_rooms:
            for person in all_people:
                assignments[(week, room, person)] = model.NewBoolVar(f'assignment_W{week}_R{room}_P{person}')

    # Constraints
    for week in all_weeks:
        # the room requirements
        model.AddExactlyOne(
            [assignments[week, 4, person] for person in all_people[0:4]])  # Person 0, 1, 2, 3 only in room E
        model.AddExactlyOne([assignments[week, 5, person] for person in all_people[4:6]])  # Person 4, 5 only in room F

    # each week every room is assigned to exactly one person
    for week in all_weeks:
        for room in all_rooms:
            model.AddExactlyOne([assignments[week, room, person] for person in all_people])

    # each week a person works at exactly at one room
    for person in all_people:
        for week in all_weeks:
            model.AddExactlyOne(assignments[week, room, person] for room in all_rooms)

    weight = -1
    obj_bool_vars = []
    obj_bool_coeffs = []

    obj_int_vars = []
    obj_int_coeffs = []

    for person in all_people:
        for room in all_rooms[0:6]:
            for week in all_weeks:
                for stride in range(1, num_weeks - week):
                    # doing a room more often by the same person implies a penalty
                    has_done_room1 = assignments[(week + stride, room, person)]
                    has_done_room2 = assignments[(week, room, person)]

                    penalty_var = model.NewBoolVar(f'negated_assignment_Cost{stride}{week}{room}{person}')
                    model.AddBoolAnd(penalty_var).OnlyEnforceIf([has_done_room1, has_done_room2])

                    obj_bool_vars.append(penalty_var)
                    obj_bool_coeffs.append(weight * abs(6 - stride))

    # # permit to do the same room 2 times in a row
    for person in all_people:
        for room in all_rooms[0:6]:
            for week in all_weeks[0:num_weeks-1]:
                has_done_room1 = assignments[(week + 1, room, person)]
                has_done_room2 = assignments[(week, room, person)]
                model.AddBoolOr([has_done_room1.Not(), has_done_room2.Not()])

    # # permit to do the same room with a stride of 2
    for person in all_people:
        for room in all_rooms[0:5]:
            for week in all_weeks[0:num_weeks-2]:
                has_done_room1 = assignments[(week + 2, room, person)]
                has_done_room2 = assignments[(week, room, person)]
                model.AddBoolOr([has_done_room1.Not(), has_done_room2.Not()])

    # # permit to do the same room with a stride of 3
    for person in all_people:
        for room in all_rooms[0:5]:
            for week in all_weeks[0:num_weeks-3]:
                has_done_room1 = assignments[(week + 3, room, person)]
                has_done_room2 = assignments[(week, room, person)]
                model.AddBoolOr([has_done_room1.Not(), has_done_room2.Not()])

    # # permit to do the same room with a stride of 4
    for person in all_people:
        for room in all_rooms[0:4]:
            for week in all_weeks[0:num_weeks-4]:
                has_done_room1 = assignments[(week + 4, room, person)]
                has_done_room2 = assignments[(week, room, person)]
                model.AddBoolOr([has_done_room1.Not(), has_done_room2.Not()])


    # start with a certain combination of people
    model.AddBoolAnd(assignments[(0, 0, 3)])
    model.AddBoolAnd(assignments[(0, 1, 1)])
    model.AddBoolAnd(assignments[(0, 2, 2)])
    model.AddBoolAnd(assignments[(0, 3, 4)])
    model.AddBoolAnd(assignments[(0, 4, 0)])
    model.AddBoolAnd(assignments[(0, 5, 5)])

    model.Maximize(sum(obj_bool_vars[i] * obj_bool_coeffs[i] for i in range(len(obj_bool_vars)))
                   + sum(obj_int_vars[i] * obj_int_coeffs[i] for i in range(len(obj_int_coeffs))))

    # Solve the model
    print("Solving ...")
    solver = cp_model.CpSolver()
    # solver.parameters.max_time_in_seconds = 10
    # only search for 10 solution, not all
    #solution_printer = VarArraySolutionPrinterWithLimit(10)
    #status = solver.Solve(model, solution_printer)
    status = solver.Solve(model)

    print("  - wall time       : %f s" % solver.WallTime())

    counter_dict = {}

    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        column_per_room = {}
        for room in all_rooms:
            column_per_room[room] = []

        counter_per_person = {}
        for person in all_people:
            counter_per_person[person] = 0

        df = DataFrame()
        for room in all_rooms:
            for week in all_weeks:
                for person in all_people:
                    isAssigned = solver.Value(assignments[(week, room, person)])
                    if isAssigned != 0:
                        counter_per_person[person] += 1
                        column_per_room[room].append(name_persons[person])
        for room in all_rooms:
            for week in all_weeks:
                for person in all_people:
                    isAssigned = solver.Value(assignments[(week, room, person)])
                    if isAssigned != 0:
                        df[name_rooms[room]] = Series(column_per_room[room])
        df.to_excel('scheduling_CP.xlsx', sheet_name='scheduling', index=True)

        print(counter_per_person)

        for room in all_rooms:
            for person in all_people:
                counter_dict[room, person] = 0
                for week in all_weeks:
                    isAssigned = solver.Value(assignments[(week, room, person)])
                    if isAssigned != 0:
                        counter_dict[room, person] += 1
                print(f" Room {room} By Person {person} in week for: {counter_dict[room, person]} times")
    else:
        print("No feasible solution found.")

if __name__ == "__main__":
    main()
