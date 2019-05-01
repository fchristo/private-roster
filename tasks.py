from roster import Roster
from invoke import task
import ast

"""This is the invoke tasks file, made for ease of use. Call it using invoke (task_name)"""


@task
def get_student_names(ctx, filename="Jones_2019.xlsx"):
    with Roster(filename) as r:
        print(r.get_student_names())


@task
def get_student(ctx, student_identifier, filename="Jones_2019.xlsx"):
    with Roster(filename) as r:
        if is_intstring(student_identifier):
            print(r.get_student(int(student_identifier)))
        else:
            print(r.get_student(student_identifier))


@task
def delete_student(ctx, student_identifier, filename="Jones_2019.xlsx"):
    with Roster(filename) as r:
        if is_intstring(student_identifier):
            r.delete_student(int(student_identifier))
        else:
            r.delete_student(student_identifier)

        updated_file = filename.split(".xlsx")[0] + "_Updated.xlsx"
        print("Student deleted. See " + updated_file + " to review sheets.")


@task
def class_average(ctx, filename="Jones_2019.xlsx"):
    with Roster(filename) as r:
        print(r.class_average())


@task
def add_grades(ctx, student_dict, filename="Jones_2019.xlsx"):
    with Roster(filename) as r:
        data = ast.literal_eval(student_dict)
        print(data)
        r.add_grades(data)
        updated_file = filename.split(".xlsx")[0] + "_Updated.xlsx"
        print(
            "Grades added. See "
            + updated_file
            + " or run invoke get_student(<student>) to review changes."
        )


@task
def gen_docs(ctx):
    ctx.run("sphinx-build -b html ./ _build")


def is_intstring(input_arg: str):
    try:
        int(input_arg)
        return True
    except ValueError:
        return False
