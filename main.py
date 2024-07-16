import boto3
import uuid
import json

dynamodb = boto3.resource('dynamodb', endpoint_url='http://localhost:8000')

def delete_data(table_name):
    table = dynamodb.Table(table_name)
    key = 'id' if table_name != 'subject' else 'code'
    id_delete = input(f"Delete {table_name} by {key}: ").strip()
    if id_delete == 'back':
        return

    response = table.get_item(Key={key: id_delete})
    if 'Item' not in response:
        print(f"{table_name.capitalize()} {key} {id_delete} not found.")
    else:
        table.delete_item(Key={key: id_delete})
        print(f"{table_name.capitalize()} {key} {id_delete} deleted.")

def create_university():
    table = dynamodb.Table('university')
    name = input("Name: ").strip()
    if name == 'back': return
    address = input("Address: ").strip()
    if address == 'back': return

    table.put_item(Item={'id': str(uuid.uuid4()), 'name': name, 'address': address})
    print(f"University {name} created.")

def read_university():
    table = dynamodb.Table('university')
    mode = input("Read university (find, list): ").strip().lower()
    if mode == 'find':
        university_id = input("University ID: ").strip()
        response = table.get_item(Key={'id': university_id})
        if 'Item' in response:
            print(f"ID: {response['Item']['id']}, Name: {response['Item']['name']}, Address: {response['Item']['address']}")
        else:
            print("University not found.")
    elif mode == 'list':
        response = table.scan()
        if 'Items' in response and response['Items']:
            for uni in response['Items']:
                print(f"ID: {uni['id']}, Name: {uni['name']}, Address: {uni['address']}")
        else:
            print("No universities found.")

def update_university():
    table = dynamodb.Table('university')
    university_id = input("University ID: ").strip()
    if university_id == 'back': return

    response = table.get_item(Key={'id': university_id})
    if 'Item' in response:
        mode_update = input("Update (name, address): ").strip().lower()
        new_value = input(f"New {mode_update}: ").strip()
        if new_value == 'back': return

        table.update_item(
            Key={'id': university_id},
            UpdateExpression=f"SET #attr = :val",
            ExpressionAttributeNames={'#attr': mode_update},
            ExpressionAttributeValues={':val': new_value}
        )
        print(f"University {mode_update} updated.")
    else:
        print("University not found.")

def manage_university():
    while True:
        action = input("Manage university (create, read, update, delete, back): ").strip().lower()
        if action == 'create':
            create_university()
        elif action == 'read':
            read_university()
        elif action == 'update':
            update_university()
        elif action == 'delete':
            delete_data('university')
        elif action == 'back':
            break

def create_subject():
    table = dynamodb.Table('subject')
    code = input("Code: ").strip()
    if code == 'back': return
    name = input("Name: ").strip()
    if name == 'back': return
    limit = input("Limit: ").strip()
    if limit == 'back': return
    field = input("Field: ").strip()
    if limit == 'back': return

    table.put_item(Item={'code': code, 'name': name, 'limit': int(limit), 'students': [], 'can_enroll_field': field})
    print(f"Subject {name} created.")

def read_subject():
    table = dynamodb.Table('subject')
    mode = input("Read subject (find, list): ").strip().lower()
    if mode == 'find':
        subject_code = input("Subject Code: ").strip()
        response = table.get_item(Key={'code': subject_code})
        if 'Item' in response:
            print(f"Code: {response['Item']['code']}, Name: {response['Item']['name']}, Limit: {response['Item']['limit']}")
        else:
            print("Subject not found.")
    elif mode == 'list':
        response = table.scan()
        if 'Items' in response and response['Items']:
            for sub in response['Items']:
                print(f"Code: {sub['code']}, Name: {sub['name']}, Limit: {sub['limit']}")
        else:
            print("No subjects found.")

def update_student():
    table = dynamodb.Table('student')
    student_id = input("Student ID: ").strip()
    if student_id == 'back':
        return

    response = table.get_item(Key={'id': student_id})
    if 'Item' in response:
        student = response['Item']
        mode_update = input("Update (name, lastname, age, grade, field, subject, university): ").strip().lower()
        new_value = input(f"New {mode_update}: ").strip()
        if new_value == 'back':
            return

        if mode_update == 'subject':
            subject_code = new_value
            subject_response = dynamodb.Table('subject').get_item(Key={'code': subject_code})
            if 'Item' not in subject_response:
                print("Subject not found.")
                return

            subject = subject_response['Item']

            if subject['can_enroll_field'] is not None and student['field'] != subject['can_enroll_field']:
                print(f"Student cannot enroll in this subject as their field ({student['field']}) does not match the required field ({subject['can_enroll_field']}).")
                return

            subject_limit = int(subject['limit']) if 'limit' in subject else float('inf')

            if len(subject['students']) >= subject_limit:
                print(f"Subject {subject_code} has reached its enrollment limit.")
                return

            if 'enroll_subject' not in student:
                student['enroll_subject'] = []

            enroll_subject_list = json.loads(student['enroll_subject']) if student['enroll_subject'] else []
            if subject_code not in enroll_subject_list:
                enroll_subject_list.append(subject_code)
                student['enroll_subject'] = json.dumps(enroll_subject_list)

            table.update_item(
                Key={'id': student_id},
                UpdateExpression="SET #enroll = :val",
                ExpressionAttributeNames={'#enroll': 'enroll_subject'},
                ExpressionAttributeValues={':val': student['enroll_subject']}
            )
            print(f"Student enrolled in subject {subject_code}.")

            # Update subject table with student data, avoid duplicates
            student_data = {
                'id': student['id'],
                'name': student['name'],
                'lastname': student['lastname'],
                'age': student['age'],
                'grade': student['grade']
            }

            subject_table = dynamodb.Table('subject')
            subject_response = subject_table.get_item(Key={'code': subject_code})
            if 'Item' in subject_response:
                subject = subject_response['Item']
                if 'students' not in subject:
                    subject['students'] = []

                existing_students = subject['students']
                if student_data not in existing_students:
                    existing_students.append(student_data)
                    subject_table.update_item(
                        Key={'code': subject_code},
                        UpdateExpression="SET #students = :val",
                        ExpressionAttributeNames={'#students': 'students'},
                        ExpressionAttributeValues={':val': existing_students}
                    )
                    print(f"Subject {subject_code} updated with student data.")
                else:
                    print(f"Student {student['id']} already enrolled in subject {subject_code}.")
            else:
                print("Subject not found.")
        elif mode_update == 'university':
            university_id = new_value
            university_response = dynamodb.Table('university').get_item(Key={'id': university_id})
            if 'Item' not in university_response:
                print("University not found.")
                return

            university = university_response['Item']
            table.update_item(
                Key={'id': student_id},
                UpdateExpression=f"SET #attr = :val",
                ExpressionAttributeNames={'#attr': 'university'},
                ExpressionAttributeValues={':val': json.dumps(university)}
            )
            print(f"Student enrolled in university {university['name']}.")
        else:
            table.update_item(
                Key={'id': student_id},
                UpdateExpression=f"SET #attr = :val",
                ExpressionAttributeNames={'#attr': mode_update},
                ExpressionAttributeValues={':val': new_value}
            )
            print(f"Student {mode_update} updated.")
    else:
        print("Student not found.")

def manage_subject():
    while True:
        action = input("Manage subject (create, read, update, delete, back): ").strip().lower()
        if action == 'create':
            create_subject()
        elif action == 'read':
            read_subject()
        elif action == 'update':
            update_subject()
        elif action == 'delete':
            delete_data('subject')
        elif action == 'back':
            break

def create_student():
    table = dynamodb.Table('student')
    student_id = str(uuid.uuid4())
    firstname = input("First name: ").strip()
    if firstname == 'back': return
    lastname = input("Last name: ").strip()
    if lastname == 'back': return
    age = input("Age: ").strip()
    if age == 'back': return
    grade = input("Grade: ").strip()
    if grade == 'back': return
    field = input("Field: ").strip()
    if field == 'back': return

    table.put_item(Item={'id': student_id, 'name': firstname, 'lastname': lastname, 'age': age, 'grade': grade, 'field': field, 'university': None, 'enroll_subject': None})
    print(f"Student {firstname} {lastname} created.")

def read_student():
    table = dynamodb.Table('student')
    mode = input("Read student (find, list): ").strip().lower()
    if mode == 'find':
        student_id = input("Student ID: ").strip()
        response = table.get_item(Key={'id': student_id})
        if 'Item' in response:
            student = response['Item']
            print(f"ID: {student['id']}, Name: {student['name']} {student['lastname']}, Age: {student['age']}, Grade: {student['grade']}, Field: {student['field']}, University: {student['university']}, Enrolled Subjects: {student['enroll_subject']}")
        else:
            print("Student not found.")
    elif mode == 'list':
        response = table.scan()
        if 'Items' in response and response['Items']:
            for student in response['Items']:
                print(f"ID: {student['id']}, Name: {student['name']} {student['lastname']}, Age: {student['age']}, Grade: {student['grade']}, Field: {student['field']}, University: {student['university']}, Enrolled Subjects: {student['enroll_subject']}")
        else:
            print("No students found.")
def update_subject():
    table = dynamodb.Table('subject')
    subject_code = input("Subject Code: ").strip()
    if subject_code == 'back': return

    response = table.get_item(Key={'code': subject_code})
    if 'Item' in response:
        mode = ['name','limit']
        mode_update = input("Update (name, limit, field): ").strip().lower()
        new_value = input(f"New {mode_update}: ").strip()
        if new_value == 'back': return
        if mode_update in mode:
            table.update_item(
                Key={'code': subject_code},
                UpdateExpression=f"SET #attr = :val",
                ExpressionAttributeNames={'#attr': mode_update},
                ExpressionAttributeValues={':val': new_value}
            )
        elif mode_update == 'field':
            table.update_item(
                Key={'code': subject_code},
                UpdateExpression=f"SET #attr = :val",
                ExpressionAttributeNames={'#attr': 'can_enroll_field'},
                ExpressionAttributeValues={':val': new_value}
            )
        print(f"Subject {mode_update} updated.")
    else:
        print("Subject not found.")

def update_student():
    table = dynamodb.Table('student')
    student_id = input("Student ID: ").strip()
    if student_id == 'back':
        return

    response = table.get_item(Key={'id': student_id})
    if 'Item' in response:
        student = response['Item']
        mode_update = input("Update (name, lastname, age, grade, field, subject, university): ").strip().lower()
        new_value = input(f"New {mode_update}: ").strip()
        if new_value == 'back':
            return

        if mode_update == 'subject':
            subject_code = new_value
            subject_response = dynamodb.Table('subject').get_item(Key={'code': subject_code})
            if 'Item' not in subject_response:
                print("Subject not found.")
                return

            subject = subject_response['Item']

            if subject['can_enroll_field'] is not None and student['field'] != subject['can_enroll_field']:
                print(f"Student cannot enroll in this subject as their field ({student['field']}) does not match the required field ({subject['can_enroll_field']}).")
                return

            subject_limit = int(subject['limit']) if 'limit' in subject else float('inf')

            if len(subject['students']) >= subject_limit:
                print(f"Subject {subject_code} has reached its enrollment limit.")
                return

            if 'enroll_subject' not in student:
                student['enroll_subject'] = []

            enroll_subject_list = json.loads(student['enroll_subject']) if student['enroll_subject'] else []
            if subject_code not in enroll_subject_list:
                enroll_subject_list.append(subject_code)
                student['enroll_subject'] = json.dumps(enroll_subject_list)

            table.update_item(
                Key={'id': student_id},
                UpdateExpression="SET #enroll = :val",
                ExpressionAttributeNames={'#enroll': 'enroll_subject'},
                ExpressionAttributeValues={':val': student['enroll_subject']}
            )
            print(f"Student enrolled in subject {subject_code}.")

            student_data = {
                'id': student['id'],
                'name': student['name'],
                'lastname': student['lastname'],
                'age': student['age'],
                'grade': student['grade']
            }

            subject_table = dynamodb.Table('subject')
            subject_response = subject_table.get_item(Key={'code': subject_code})
            if 'Item' in subject_response:
                subject = subject_response['Item']
                if 'students' not in subject:
                    subject['students'] = []

                existing_students = subject['students']
                if student_data not in existing_students:
                    existing_students.append(student_data)
                    subject_table.update_item(
                        Key={'code': subject_code},
                        UpdateExpression="SET #students = :val",
                        ExpressionAttributeNames={'#students': 'students'},
                        ExpressionAttributeValues={':val': existing_students}
                    )
                    print(f"Subject {subject_code} updated with student data.")
                else:
                    print(f"Student {student['id']} already enrolled in subject {subject_code}.")
            else:
                print("Subject not found.")
        elif mode_update == 'university':
            university_id = new_value
            university_response = dynamodb.Table('university').get_item(Key={'id': university_id})
            if 'Item' not in university_response:
                print("University not found.")
                return

            university = university_response['Item']
            table.update_item(
                Key={'id': student_id},
                UpdateExpression=f"SET #attr = :val",
                ExpressionAttributeNames={'#attr': 'university'},
                ExpressionAttributeValues={':val': json.dumps(university)}
            )
            print(f"Student enrolled in university {university['name']}.")
        else:
            table.update_item(
                Key={'id': student_id},
                UpdateExpression=f"SET #attr = :val",
                ExpressionAttributeNames={'#attr': mode_update},
                ExpressionAttributeValues={':val': new_value}
            )
            print(f"Student {mode_update} updated.")
    else:
        print("Student not found.")
    table = dynamodb.Table('student')
    student_id = input("Student ID: ").strip()
    if student_id == 'back': return

    response = table.get_item(Key={'id': student_id})
    if 'Item' in response:
        student = response['Item']
        mode_update = input("Update (name, lastname, age, grade, field, subject, university): ").strip().lower()
        new_value = input(f"New {mode_update}: ").strip()
        if new_value == 'back': return

        if mode_update == 'subject':
            subject_code = new_value
            subject_response = dynamodb.Table('subject').get_item(Key={'code': subject_code})
            if 'Item' not in subject_response:
                print("Subject not found.")
                return

            subject = subject_response['Item']
            
            # Convert subject limit to int for comparison
            subject_limit = int(subject['limit']) if 'limit' in subject else float('inf')

            # Check if subject enrollment limit is reached
            if len(subject['students']) >= subject_limit:
                print(f"Subject {subject_code} has reached its enrollment limit.")
                return

            if subject['can_enroll_field'] is not None and student['field'] != subject['can_enroll_field']:
                print(f"Student cannot enroll in this subject as their field ({student['field']}) does not match the required field ({subject['can_enroll_field']}).")
                return

            # Update student table: append new subject to enroll_subject list
            if 'enroll_subject' not in student:
                student['enroll_subject'] = []

            enroll_subject_list = json.loads(student['enroll_subject']) if student['enroll_subject'] else []
            if subject_code not in enroll_subject_list:
                enroll_subject_list.append(subject_code)
                student['enroll_subject'] = json.dumps(enroll_subject_list)

            table.update_item(
                Key={'id': student_id},
                UpdateExpression="SET #enroll = :val",
                ExpressionAttributeNames={'#enroll': 'enroll_subject'},
                ExpressionAttributeValues={':val': student['enroll_subject']}
            )
            print(f"Student enrolled in subject {subject_code}.")

            # Update subject table with student data, avoid duplicates
            student_data = {
                'id': student['id'],
                'name': student['name'],
                'lastname': student['lastname'],
                'age': student['age'],
                'grade': student['grade']
            }

            subject_table = dynamodb.Table('subject')
            subject_response = subject_table.get_item(Key={'code': subject_code})
            if 'Item' in subject_response:
                subject = subject_response['Item']
                if 'students' not in subject:
                    subject['students'] = []

                existing_students = subject['students']
                if student_data not in existing_students:
                    existing_students.append(student_data)
                    subject_table.update_item(
                        Key={'code': subject_code},
                        UpdateExpression="SET #students = :val",
                        ExpressionAttributeNames={'#students': 'students'},
                        ExpressionAttributeValues={':val': existing_students}
                    )
                    print(f"Subject {subject_code} updated with student data.")
                else:
                    print(f"Student {student['id']} already enrolled in subject {subject_code}.")
            else:
                print("Subject not found.")
        elif mode_update == 'university':
            university_id = new_value
            university_response = dynamodb.Table('university').get_item(Key={'id': university_id})
            if 'Item' not in university_response:
                print("University not found.")
                return

            university = university_response['Item']
            table.update_item(
                Key={'id': student_id},
                UpdateExpression=f"SET #attr = :val",
                ExpressionAttributeNames={'#attr': 'university'},
                ExpressionAttributeValues={':val': json.dumps(university)}
            )
            print(f"Student enrolled in university {university['name']}.")
        else:
            table.update_item(
                Key={'id': student_id},
                UpdateExpression=f"SET #attr = :val",
                ExpressionAttributeNames={'#attr': mode_update},
                ExpressionAttributeValues={':val': new_value}
            )
            print(f"Student {mode_update} updated.")
    else:
        print("Student not found.")


def manage_student():
    while True:
        action = input("Manage student (create, read, update, delete, back): ").strip().lower()
        if action == 'create':
            create_student()
        elif action == 'read':
            read_student()
        elif action == 'update':
            update_student()
        elif action == 'delete':
            delete_data('student')
        elif action == 'back':
            break

def main():
    while True:
        entity = input("Choose entity (university, student, subject, exit): ").strip().lower()
        if entity == 'university':
            manage_university()
        elif entity == 'student':
            manage_student()
        elif entity == 'subject':
            manage_subject()
        elif entity == 'exit':
            break

if __name__ == "__main__":
    main()
