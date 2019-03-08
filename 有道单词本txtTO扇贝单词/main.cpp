#include <iostream>
#include<string>
#include<fstream>

using namespace std;

bool isNumber(string str,int i);
string extractWord(string str);

int main() {
	string fileName = "有道单词本.txt";
	cout << "输入文件名：";
	cin >> fileName;
	fstream in(fileName, ios::in||ios::out);
	fstream out;
	//cout << in.is_open();
	string temp;
	//char buf[80];
	int count = 0;//扇贝单词每次只能添加10个单词，每个单词一行
	int txtNumber = 1;
	while (getline(in,temp)) {
		//cout << temp << endl;
		if (isNumber(temp,0)) {
			if (count == 10) {
				count = 0; txtNumber++;
			}
			if (count == 0) {
				out.close();//给忘啦！
				string tempFileName = to_string(txtNumber)+".txt";
				out.open(tempFileName, ios::out);
				cout << out.is_open() << endl << endl << endl;
			}
			string result = extractWord(temp) + '\n';
			char * p = (char*)result.data();
			p[result.length() + 1] = '\n';
			cout<<p;
			out << p;
			++count;
		}
		
	}
	system("pause");
	return 0;
}

bool isNumber(string str,int i) {
	char first = str[i];
	if (first == '1' ||
		first == '2' ||
		first == '3' ||
		first == '4' ||
		first == '5' ||
		first == '6' ||
		first == '7' ||
		first == '8' ||
		first == '9' ||
		first == '0') {
		return true;
	}
	return false;
}

string extractWord(string str) {
	char word[100];
	int m = 0;
	bool flag = false;
	for (int i = 0; i < str.length(); ++i) {
		if (!flag) {
			if (str[i] == ' ')flag = true;
		}
		else {
			if (str[i] == ' ') {
				word[m] = '\0';
				string result(word);
				return result;
			}
			word[m] = str[i];
			++m;
		}
	}
}