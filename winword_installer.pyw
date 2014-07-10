#coding: utf-8 -*-
from doccleaner.plugins.winword import wordaddin

def main():
	wordaddin.main(['--register'])

if __name__ == '__main__':
	main()