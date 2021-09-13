from django.shortcuts import render, redirect
#from .forms import Tes
from django.contrib.auth.decorators import login_required
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework.mixins import UpdateModelMixin, DestroyModelMixin

from report.models import Repair
from report.serializers import RepairSerializer

'''
from .models import Repair
from django.views.decorators.csrf import csrf_exempt

from django.http.response import JsonResponse
from rest_framework.parsers import JSONParser 
from rest_framework import status
 
from report.models import Repair
from report.serializers import RepairSerializer
from rest_framework.views import APIView
#from .tasks import send_mail_task
# Create your views here.

#@login_required
#@csrf_exempt
'''

def index(request):
    data = Repair.objects.all()
    context = {
        "repairs": data
    }
    return render(request, 'main_report.html',context)


class RepairListView(
  APIView, # Basic View class provided by the Django Rest Framework
  UpdateModelMixin, # Mixin that allows the basic APIView to handle PUT HTTP requests
  DestroyModelMixin, # Mixin that allows the basic APIView to handle DELETE HTTP requests
):

  def get(self, request, id=None):
    if id:
      # If an id is provided in the GET request, retrieve the Repair item by that id
      try:
        # Check if the Repair item the user wants to update exists
        queryset = Repair.objects.get(id=id)
      except Repair.DoesNotExist:
        # If the Repair item does not exist, return an error response
        return Response({'errors': 'This Repair item does not exist.'}, status=400)

      # Serialize Repair item from Django queryset object to JSON formatted data
      read_serializer = RepairSerializer(queryset)

    else:
      # Get all Repair items from the database using Django's model ORM
      queryset = Repair.objects.all()

      # Serialize list of Repairs item from Django queryset object to JSON formatted data
      read_serializer = RepairSerializer(queryset, many=True)

    # Return a HTTP response object with the list of Repair items as JSON
    return Response(read_serializer.data)


  def post(self, request):
    # Pass JSON data from user POST request to serializer for validation
    create_serializer = RepairSerializer(data=request.data)

    # Check if user POST data passes validation checks from serializer
    if create_serializer.is_valid():

      # If user data is valid, create a new Repair item record in the database
      Repair_item_object = create_serializer.save()

      # Serialize the new Repair item from a Python object to JSON format
      read_serializer = RepairSerializer(Repair_item_object)

      # Return a HTTP response with the newly created Repair item data
      return Response(read_serializer.data, status=201)

    # If the users POST data is not valid, return a 400 response with an error message
    return Response(create_serializer.errors, status=400)


  def put(self, request, id=None):
    try:
      # Check if the Repair item the user wants to update exists
      Repair_item = Repair.objects.get(id=id)
    except Repair.DoesNotExist:
      # If the Repair item does not exist, return an error response
      return Response({'errors': 'This Repair item does not exist.'}, status=400)

    # If the Repair item does exists, use the serializer to validate the updated data
    update_serializer = RepairSerializer(Repair_item, data=request.data)

    # If the data to update the Repair item is valid, proceed to saving data to the database
    if update_serializer.is_valid():

      # Data was valid, update the Repair item in the database
      Repair_item_object = update_serializer.save()

      # Serialize the Repair item from Python object to JSON format
      read_serializer = RepairSerializer(Repair_item_object)

      # Return a HTTP response with the newly updated Repair item
      return Response(read_serializer.data, status=200)

    # If the update data is not valid, return an error response
    return Response(update_serializer.errors, status=400)

    '''
  def delete(self, request, id=None):
    try:
      # Check if the Repair item the user wants to update exists
      Repair_item = Repair.objects.get(id=id)
    except Repair.DoesNotExist:
      # If the Repair item does not exist, return an error response
      return Response({'errors': 'This Repair item does not exist.'}, status=400)

    # Delete the chosen Repair item from the database
    Repair_item.delete()

    # Return a HTTP response notifying that the Repair item was successfully deleted
    return Response(status=204)
    '''
def success(request):
    return render(request, 'success_report.html')
    