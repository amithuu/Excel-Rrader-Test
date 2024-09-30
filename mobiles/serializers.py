from rest_framework import serializers

class DynamicExcelDataSerializer(serializers.Serializer):
    data = serializers.DictField(required=True)