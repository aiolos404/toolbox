#include <iostream>
#include <fstream>
 
using namespace std;
 
int main()
{
    struct point
{
  int x;
  int y; 
};

ifstream infile("a.txt");
int a, b;
point pnts[30];
int i(0);
 while(infile >> a >> b)
{ 
printf("%d\n",i);
   pnts[i].x = a;
   printf("%d\n",pnts[i].x);
   pnts[i].y = b;
   printf("%d\n",pnts[i].y);
	i++;
	
}
}