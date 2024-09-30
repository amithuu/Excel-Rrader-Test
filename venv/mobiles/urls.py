from django.urls import path
from .views import fetch_excel_data, create_new_column, rename_columns, update_values, empty_excel

urlpatterns = [
    path('data/', fetch_excel_data, name='fetch_excel_data'),
    path('data/create-column/', create_new_column, name='create_new_column'),
    path('data/add_row/', rename_columns, name='rename_columns'),
    path('data/update/', update_values, name='update_values'),
    path('data/empty/', empty_excel, name='empty_excel'),
]
