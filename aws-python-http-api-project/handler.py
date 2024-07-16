import boto3
import uuid
import json
from decimal import Decimal
from datetime import datetime,timezone

dynamodb = boto3.resource('dynamodb', endpoint_url='http://localhost:8000')


def delete_data(event, context):
    table_name = event['pathParameters']['table_name']
    table = dynamodb.Table(table_name)
    key = 'id' if table_name != 'subject' else 'code'
    body = json.loads(event['body'])
    id_delete = body.get(key)
    if not id_delete:
        return {
            "statusCode": 400,
            "body": json.dumps({"message": f"Missing {key} in request body"})
        }

    response = table.get_item(Key={key: id_delete})
    if 'Item' not in response:
        return {
            "statusCode": 404,
            "body": json.dumps({"message": f"{table_name.capitalize()} {key} {id_delete} not found"})
        }
    else:
        table.delete_item(Key={key: id_delete})
        return {
            "statusCode": 200,
            "body": json.dumps({"message": f"{table_name.capitalize()} {key} {id_delete} deleted"})
        }


def decimal_to_float(obj):
    if isinstance(obj, list):
        for i in range(len(obj)):
            obj[i] = decimal_to_float(obj[i])
    elif isinstance(obj, dict):
        for k, v in obj.items():
            obj[k] = decimal_to_float(v)
    elif isinstance(obj, Decimal):
        if obj % 1 == 0:
            obj = int(obj)
        else:
            obj = float(obj)
    return obj


def list_data(event, context):
    table_name = event['pathParameters']['table_name']
    table = dynamodb.Table(table_name)

    response = table.scan()
    if 'Items' in response:
        items = decimal_to_float(response['Items'])
        return {
            "statusCode": 200,
            "body": json.dumps(items)
        }
    else:
        return {
            "statusCode": 404,
            "body": json.dumps({"message": "No items found in the table"})
        }


def list_one_data(event, context):
    table_name = event['pathParameters']['table_name']
    table = dynamodb.Table(table_name)
    if table_name == 'subject':
        print(table_name)
        primary_key = event['queryStringParameters'].get('code')
        response = table.get_item(Key={'code': primary_key})
    else:
        primary_key = event['queryStringParameters'].get('id')
        response = table.get_item(Key={'id': primary_key})

    if not primary_key:
        return {
            "statusCode": 400,
            "body": json.dumps({"message": "Primary key is required"})
        }
    if 'Item' in response:
        item = decimal_to_float(response['Item'])
        return {
            "statusCode": 200,
            "body": json.dumps(item)
        }
    else:
        return {
            "statusCode": 404,
            "body": json.dumps({"message": f"id {primary_key} not found"})
        }


def create_data(event, context):
    table_name = event['pathParameters']['table_name']
    table = dynamodb.Table(table_name)
    if table_name == 'subject':
        data = json.loads(event['body'])

        if 'name' not in data:
            data['name'] = ''
        if 'limit' not in data:
            data['limit'] = ''
        if 'can_enroll_field' not in data:
            data['can_enroll_field'] = ''
        if 'students' not in data:
            data['students'] = []


        table.put_item(Item=data)

        return {
            "statusCode": 200,
            "body": json.dumps({"item": data, "message": "Item created successfully"})
        }
    elif table_name == 'university':
        data = json.loads(event['body'])

        if 'name' not in data:
            data['name'] = ''
        if 'address' not in data:
            data['address'] = ''
        if 'id' not in data:
            data['id'] = str(uuid.uuid4())

        table.put_item(Item=data)
        return {
            "statusCode": 200,
            "body": json.dumps({"item": data, "message": "Item created successfully"})
        }
    elif table_name == 'student':
        data = json.loads(event['body'])
        if 'id' not in data:
            data['id'] = str(uuid.uuid4())
        if 'enroll_subject' not in data:
            data['enroll_subject'] = []
        if 'field' not in data:
            data['field'] = ''
        if 'university' not in data:
            data['university'] = {}
        if 'grade' not in data:
            data['grade'] = ''
        if 'age' not in data:
            data['age'] = ''
        if 'lastname' not in data:
            data['lastname'] = ''
        if 'name' not in data:
            data['name'] = ''
        if 'score' not in data:
            data['score'] = ''
        if 'date_create' not in data:
            data['date_create'] = datetime.now().astimezone().isoformat()
        if 'date_update' not in data:
            data['date_update'] = datetime.now().astimezone().isoformat()

        table.put_item(Item=data)

        return {
            "statusCode": 200,
            "body": json.dumps({"item": data, "message": "Item created successfully"})
        }
    elif table_name == 'life_insurance':
        data = json.loads(event['body'])
        if 'id' not in data:
            data['id'] = str(uuid.uuid4())
        table.put_item(Item=data)
        return {
            "statusCode": 200,
            "body": json.dumps({"item": data, "message": "Item created successfully"})
        }
    else:
        return {
            "statusCode": 400,
            "body": json.dumps({"message": "Invalid table name"})
        }

def update_data(event, context):
    table_name = event['pathParameters']['table_name']
    table = dynamodb.Table(table_name)
    key = 'id' if table_name != 'subject' else 'code'
    body = json.loads(event['body'])
    id = body.get(key)

    if table_name == 'university':
        return update_university(id, body)
    elif table_name == 'subject':
        return update_subject(id, body)
    elif table_name == 'student':
        return update_student(id, body)
    else:
        return {
            "statusCode": 400,
            "body": json.dumps({"message": "Invalid table name"})
        }

def update_university(id, body):
    table = dynamodb.Table('university')
    university_id = id
    response = table.get_item(Key={'id': university_id})

    if 'Item' in response:
        allowed_updates = {'name', 'address'}
        updates = {key: value for key, value in body.items() if key in allowed_updates}

        if not updates:
            return {
                "statusCode": 400,
                "body": json.dumps({"message": "No valid fields to update"})
            }

        update_expression = "SET " + ", ".join(f"#{k} = :{k}" for k in updates.keys())
        expression_attribute_names = {f"#{k}": k for k in updates.keys()}
        expression_attribute_values = {f":{k}": v for k, v in updates.items()}

        table.update_item(
            Key={'id': university_id},
            UpdateExpression=update_expression,
            ExpressionAttributeNames=expression_attribute_names,
            ExpressionAttributeValues=expression_attribute_values
        )
        return {
            "statusCode": 200,
            "body": json.dumps({"message": "University updated successfully."})
        }
    else:
        return {
            "statusCode": 400,
            "body": json.dumps({"message": "University not found."})
        }
    
def update_subject(code, body):
    table = dynamodb.Table('subject')
    response = table.get_item(Key={'code': code})

    if 'Item' in response:
        allowed_updates = {'name', 'limit', 'can_enroll_field'}
        updates = {key: value for key, value in body.items() if key in allowed_updates}

        if not updates:
            return {
                "statusCode": 400,
                "body": json.dumps({"message": "No valid fields to update"})
            }

        update_expression = "SET " + ", ".join(f"#{k} = :{k}" for k in updates.keys())
        expression_attribute_names = {f"#{k}": k for k in updates.keys()}
        expression_attribute_values = {f":{k}": v for k, v in updates.items()}

        table.update_item(
            Key={'code': code},
            UpdateExpression=update_expression,
            ExpressionAttributeNames=expression_attribute_names,
            ExpressionAttributeValues=expression_attribute_values
        )
        return {
            "statusCode": 200,
            "body": json.dumps({"message": "Subject updated successfully."})
        }
    else:
        return {
            "statusCode": 404,
            "body": json.dumps({"message": "Subject not found."})
        }
def update_student(id, body):
    table = dynamodb.Table('student')
    student_id = id
    response = table.get_item(Key={'id': student_id})

    if 'Item' in response:
        student = response['Item']
        mode_update = body.get('mode_update')
        new_value = body.get('new_value')

        # Get current timestamp
        current_timestamp = datetime.now().isoformat()

        # Update date_update in UpdateExpression and ExpressionAttributeValues
        update_expression = "SET #date_update = :timestamp"
        expression_attribute_names = {'#date_update': 'date_update'}
        expression_attribute_values = {':timestamp': current_timestamp}

        if mode_update == 'subject':
            subject_code = new_value
            subject_response = dynamodb.Table('subject').get_item(Key={'code': subject_code})

            if 'Item' not in subject_response:
                return {
                    "statusCode": 404,
                    "body": json.dumps({"message": "Subject not found"})
                }

            subject = subject_response['Item']

            if subject['can_enroll_field'] is not None and student['field'] != subject['can_enroll_field']:
                return {
                    "statusCode": 400,
                    "body": json.dumps({"message": f"Student cannot enroll in this subject as their field ({student['field']}) does not match the required field ({subject['can_enroll_field']})."})
                }

            subject_limit = int(subject['limit']) if 'limit' in subject else float('inf')

            if len(subject['students']) >= subject_limit:
                return {
                    "statusCode": 400,
                    "body": json.dumps({"message": f"Subject {subject_code} has reached its enrollment limit."})
                }

            subject_data = {
                'code': subject['code'],
                'name': subject['name'],
                'limit': subject['limit'],
                'can_enroll_field': subject['can_enroll_field']
            }

            if 'enroll_subject' not in student:
                student['enroll_subject'] = []

            existing_subjects = student['enroll_subject']
            if subject_data not in existing_subjects:
                existing_subjects.append(subject_data)

                update_expression += ", #enroll_subject = :val"
                expression_attribute_names['#enroll_subject'] = 'enroll_subject'
                expression_attribute_values[':val'] = existing_subjects

        elif mode_update == 'university':
            university_id = new_value
            university_response = dynamodb.Table('university').get_item(Key={'id': university_id})
            if 'Item' not in university_response:
                return {
                    "statusCode": 404,
                    "body": json.dumps({"message": "University not found"})
                }

            university = university_response['Item']

            update_expression += ", #university = :val"
            expression_attribute_names['#university'] = 'university'
            expression_attribute_values[':val'] = university

        else:
            update_expression += f", #{mode_update} = :val"
            expression_attribute_names[f'#{mode_update}'] = mode_update
            expression_attribute_values[':val'] = new_value

        expression_attribute_names.update(expression_attribute_names)
        expression_attribute_values.update(expression_attribute_values)

        table.update_item(
            Key={'id': student_id},
            UpdateExpression=update_expression,
            ExpressionAttributeNames=expression_attribute_names,
            ExpressionAttributeValues=expression_attribute_values
        )

        return {
            "statusCode": 200,
            "body": json.dumps({"message": f"Student {mode_update} updated."})
        }

    else:
        return {
            "statusCode": 404,
            "body": json.dumps({"message": "Student not found"})
        }
# def update_student(id, body):
#     table = dynamodb.Table('student')
#     student_id = id
#     response = table.get_item(Key={'id': student_id})

#     if 'Item' in response:
#         student = response['Item']
#         mode_update = body.get('mode_update')
#         new_value = body.get('new_value')

#         if mode_update == 'subject':
#             subject_code = new_value
#             subject_response = dynamodb.Table('subject').get_item(Key={'code': subject_code})

#             if 'Item' not in subject_response:
#                 return {
#                     "statusCode": 404,
#                     "body": json.dumps({"message": "Subject not found"})
#                 }

#             subject = subject_response['Item']

#             if subject['can_enroll_field'] is not None and student['field'] != subject['can_enroll_field']:
#                 return {
#                     "statusCode": 400,
#                     "body": json.dumps({"message": f"Student cannot enroll in this subject as their field ({student['field']}) does not match the required field ({subject['can_enroll_field']})."})
#                 }

#             subject_limit = int(subject['limit']) if 'limit' in subject else float('inf')

#             if len(subject['students']) >= subject_limit:
#                 return {
#                     "statusCode": 400,
#                     "body": json.dumps({"message": f"Subject {subject_code} has reached its enrollment limit."})
#                 }

#             subject_data = {
#                 'code': subject['code'],
#                 'name': subject['name'],
#                 'limit': subject['limit'],
#                 'can_enroll_field': subject['can_enroll_field']
#             }

#             if 'enroll_subject' not in student:
#                 student['enroll_subject'] = []

#             existing_subjects = student['enroll_subject']
#             if subject_data not in existing_subjects:
#                 existing_subjects.append(subject_data)

#                 table.update_item(
#                     Key={'id': student_id},
#                     UpdateExpression="SET #enroll_subject = :val",
#                     ExpressionAttributeNames={'#enroll_subject': 'enroll_subject'},
#                     ExpressionAttributeValues={':val': existing_subjects}
#                 )

#             return {
#                 "statusCode": 200,
#                 "body": json.dumps({"message": f"Student enrolled in subject {subject_code}."})
#             }

#         elif mode_update == 'university':
#             university_id = new_value
#             university_response = dynamodb.Table('university').get_item(Key={'id': university_id})
#             if 'Item' not in university_response:
#                 return {
#                     "statusCode": 404,
#                     "body": json.dumps({"message": "University not found"})
#                 }

#             university = university_response['Item']
#             table.update_item(
#                 Key={'id': student_id},
#                 UpdateExpression="SET #attr = :val",
#                 ExpressionAttributeNames={'#attr': 'university'},
#                 ExpressionAttributeValues={':val': university}
#             )
#             return {
#                 "statusCode": 200,
#                 "body": json.dumps({"message": f"Student enrolled in university {university['name']}."})
#             }
#         else:
#             table.update_item(
#                 Key={'id': student_id},
#                 UpdateExpression="SET #attr = :val",
#                 ExpressionAttributeNames={'#attr': mode_update},
#                 ExpressionAttributeValues={':val': new_value}
#             )
#             return {
#                 "statusCode": 200,
#                 "body": json.dumps({"message": f"Student {mode_update} updated."})
#             }
#     else:
#         return {
#             "statusCode": 404,
#             "body": json.dumps({"message": "Student not found"})
#         }
    
def delete_all_data(event, context):
    table_name = event['pathParameters']['table_name']
    table = dynamodb.Table(table_name)

    response = table.scan()
    items = response.get('Items', [])

    if not items:
        return {
            "statusCode": 404,
            "body": json.dumps({"message": "No items found in the table"})
        }
    if table_name == 'subject':
        with table.batch_writer() as batch:
            for item in items:
                batch.delete_item(Key={'code': item['code']})
    elif table_name in ['student', 'university']:
        with table.batch_writer() as batch:
            for item in items:
                batch.delete_item(Key={'id': item['id']})

    return {
        "statusCode": 200,
        "body": json.dumps({"message": f"All items deleted from {table_name} table"})
    }

