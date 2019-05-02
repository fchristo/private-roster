from roster import Roster
from invoke import task
import ast

"""This is the invoke tasks file, made for ease of use. Call it using invoke (task_name)"""


@task
def get_student_names(ctx, filename="Jones_2019.xlsx"):
    """
        Runs Roster.get_student_names()

        Parameters
        ---------
        ctx
            The context parameter required by invoke tasks.

        filename
            Name of excel file to get student names from
    """
    with Roster(filename) as r:
        print(r.get_student_names())


@task
def get_student(ctx, student_identifier, filename="Jones_2019.xlsx"):
    """
        Runs Roster.get_student_()

        Parameters
        ---------
        ctx
            The context parameter required by invoke tasks.

        student_identifier
            Either a full name as a string, or a student ID as an int

        filename
            Name of excel file to get the student
    """
    with Roster(filename) as r:
        if _is_intstring(student_identifier):
            print(r.get_student(int(student_identifier)))
        else:
            print(r.get_student(student_identifier))


@task
def delete_student(ctx, student_identifier, filename="Jones_2019.xlsx"):
    """
        Runs Roster.delete_student_()

        Parameters
        ---------
        ctx
            The context parameter required by invoke tasks.

        student_identifier
            Either a full name as a string, or a student ID as an int

        filename
            Name of excel file to delete the student from
    """
    with Roster(filename) as r:
        if _is_intstring(student_identifier):
            r.delete_student(int(student_identifier))
        else:
            r.delete_student(student_identifier)

        updated_file = filename.split(".xlsx")[0] + "_Updated.xlsx"
        print("Student deleted. See " + updated_file + " to review sheets.")


@task
def class_average(ctx, filename="Jones_2019.xlsx"):
    """
        Runs Roster.class_average()

        Parameters
        ---------
        ctx
            The context parameter required by invoke tasks.

        filename
            Name of excel file to get class average from
    """
    with Roster(filename) as r:
        print(r.class_average())


@task
def add_grades(ctx, student_dict, filename="Jones_2019.xlsx"):
    """
        Runs Roster.class_average()

        Parameters
        ---------
        ctx
            The context parameter required by invoke tasks.

        student_dict
            dictionary containing student ID and a list of assignments/grades.
                CLI Input example: invoke add_grades '{"id": 1, "grades": [(3, 90), (12, 100), (1, 10)]}'

        filename
            Name of excel file to get class average from
    """
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
    """
        Builds sphinx documentation and outputs to _build dir

        Parameters
        ---------
        ctx
            The context parameter required by invoke tasks.
    """
    ctx.run("sphinx-build -b html ./ _build")


def _is_intstring(input_arg: str):
    """
        Checks to see if the input of the command line can be an Int

        Parameters
        ---------
        input_arg
            The argument passed in by the user through the CLI
    """
    try:
        int(input_arg)
        return True
    except ValueError:
        return False
