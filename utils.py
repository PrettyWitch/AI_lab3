import math


def tra(fruit, fruit2):
    a = [0] * 9
    a[0] = fruit[2]
    if fruit[3].lower() == 'false':
        a[1] = 0
    else:
        a[1] = 1
    a[2] = int(fruit[4])
    a[3] = int(fruit[5])
    a[4] = int(fruit[6])
    a[5] = int(fruit[7])
    if fruit[8] == fruit2[8]:
        a[6] = 1
    else:
        a[6] = 0
    if fruit[9] == fruit2[9]:
        a[7] = 1
    else:
        a[7] = 0
    if fruit[10] == fruit2[10]:
        a[8] = 1
    else:
        a[8] = 0
    return a


def tra2(fruit):
    b = [0] * 9
    b[0] = fruit[2]
    if fruit[3].lower() == 'false':
        b[1] = 0
    else:
        b[1] = 1
    b[2] = int(fruit[4])
    b[3] = int(fruit[5])
    b[4] = int(fruit[6])
    b[5] = int(fruit[7])
    return b


def Euclidean(a, b):
    result = 0
    for i in range(1, len(a)):
        result += math.pow(a[i] - b[i], 2)
    result = round(math.sqrt(result), 2)
    return result


def Manhattan(a, b):
    result = 0
    for i in range(1, len(a)):
        result += abs(a[i] - b[i])
    return result


# def Minkowski(a, b):
#     result = 0
#     p = 1.5
#     for i in range(1, len(a)):
#         result += math.pow(a[i] - b[i], p)
#     result = math.pow(result, 1 / p)
#     result = round(result, 2)
#     return result

def Cosine(a, b):
    temp1 = 1
    temp2 = 1
    temp3 = 1
    for i in range(1, len(a)):
        temp1 += a[i] * b[i]
        temp2 += math.pow(a[i], 2)
        temp3 += math.pow(a[i], 2)

    result = temp1 / (math.sqrt(temp2) * math.sqrt(temp3))
    result = round(result, 2)
    return result


def split(a):
    array = []
    for i in str(a):
        array.append(i)
    return array


def disNode(a, b):
    arr1 = split(a)
    arr2 = split(b)
    temp = 0
    for i in range(len(arr1)):
        if arr1[i] == arr2[i]:
            temp += 1
        else:
            break
    affinity = len(arr1) + len(arr2) - temp * 2 - 1
    return affinity


def PearsonCorrelation(xData, yData):
    if len(xData) != len(yData):
        raise RuntimeError('Incorrect data')
    # 拿到两个数据的平均值
    xMeans = getMeans(xData)
    yMeans = getMeans(yData)
    # 计算皮尔逊系数的分子
    numerator = generateNumerator(xData, xMeans, yData, yMeans)
    # 计算皮尔逊系数的分母
    denominator = generateDenomiator(xData, xMeans, yData, yMeans)
    # 计算皮尔逊系数
    result = round(numerator / denominator, 4)
    return result


# 分子
def generateNumerator(xData, xMeans, yData, yMeans):
    numerator = 0
    for i in range(1, len(xData)):
        numerator += (xData[i] - xMeans) * (yData[i] - yMeans)
    return numerator


# 分母
def generateDenomiator(xData, xMeans, yData, yMeans):
    xSum = 0
    for i in range(1, len(xData)):
        xSum += (xData[i] - xMeans) * (xData[i] - xMeans)
    ySum = 0
    for i in range(1, len(yData)):
        ySum += (yData[i] - yMeans) * (yData[i] - yMeans)
    return math.sqrt(xSum) * math.sqrt(ySum)


# 平均值
def getMeans(datas):
    sum = 0
    for i in range(1, len(datas)):
        sum += datas[i]
    mean = sum / (len(datas) - 1)
    return mean

