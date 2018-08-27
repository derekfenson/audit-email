
from abc import ABC, abstractmethod


class Sender(ABC):
	"""
	Abstract class for creating and sending emails based on audit report from Auditor class
	"""
	def __init__(self):
		super().__init__()

	
