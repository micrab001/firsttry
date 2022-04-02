def chk_cell(i, j):
    global matrix_chk, matrix, matrix_sqr
    if matrix[i][j] == matrix[i][j - 1]:
        matrix_sqr[i][j] = matrix[i][j]
    if matrix[i][j] > matrix[i][j - 1]:
        matrix_chk[i][j] = False

    if matrix[i][j] == matrix[i + 1][j]:
        matrix_sqr[i][j] = matrix[i][j]
    if matrix[i][j] > matrix[i+1][j]:
        matrix_chk[i][j] = False

    if matrix[i][j] == matrix[i][j + 1]:
        matrix_sqr[i][j] = matrix[i][j]
    if matrix[i][j] > matrix[i][j + 1]:
        matrix_chk[i][j] = False

    if matrix[i][j] == matrix[i - 1][j]:
        matrix_sqr[i][j] = matrix[i][j]
    if matrix[i][j] > matrix[i - 1][j]:
        matrix_chk[i][j] = False

# def chk_sqr(i, j):
#     global loc_sqr, matrix_sqr
#     loc_sqr.append((i,j))
#     tmp = matrix_sqr[i][j]
#     matrix_sqr[i][j] = "-"
#     if tmp == matrix_sqr[i - 1][j]:
#         chk_sqr(i - 1, j)
#     if tmp == matrix_sqr[i][j + 1]:
#         chk_sqr(i, j + 1)
#     if tmp == matrix_sqr[i +1][j]:
#         chk_sqr(i + 1, j)
#     if tmp == matrix_sqr[i][j-1]:
#         chk_sqr(i, j - 1)

def chk_sqr(i, j):
    global loc_sqr, matrix_sqr
    loc_sqr.append((i,j, matrix_sqr[i][j]))
    matrix_sqr[i][j] = "-"
    last_col_el = 1
    while last_col_el != 0:
        last_col_el = 0
        for el in loc_sqr[(-1)*last_col_el:]:
            i = el[0]
            j = el[1]
            if el[2] == matrix_sqr[i - 1][j]:
                loc_sqr.append((i -1, j, matrix_sqr[i - 1][j]))
                matrix_sqr[i-1][j] = "-"
                last_col_el += 1
            if el[2] == matrix_sqr[i][j + 1]:
                loc_sqr.append((i, j+1, matrix_sqr[i][j + 1]))
                matrix_sqr[i][j+1] = "-"
                last_col_el += 1
            if el[2] == matrix_sqr[i +1][j]:
                loc_sqr.append((i+1, j, matrix_sqr[i + 1][j]))
                matrix_sqr[i+1][j] = "-"
                last_col_el += 1
            if el[2] == matrix_sqr[i][j-1]:
                loc_sqr.append((i, j-1, matrix_sqr[i][j-1]))
                matrix_sqr[i][j-1] = "-"
                last_col_el += 1


n_str, m_col = map(int, input().split())
matrix = [[10001 for j in range(m_col + 2)] for i in range(n_str + 2)]
for i in range(1, n_str+1):
    tmp = input().split()
    for j in range(1, m_col+1):
        matrix[i][j] = int(tmp[j-1])

matrix_chk = [[True for j in range(m_col + 2)] for i in range(n_str + 2)]
matrix_sqr = [["-" for j in range(m_col + 2)] for i in range(n_str + 2)]

for i in range(1, n_str + 1):
    for j in range(1, m_col + 1):
        chk_cell(i, j)

for i in range(1,n_str+1):
    for j in range(1, m_col+1):
        if matrix_sqr[i][j] != "-":
            loc_sqr = []
            chk_sqr(i, j)
            if all([matrix_chk[el[0]][el[1]] for el in loc_sqr]):
                for ii in range(len(loc_sqr)):
                    if ii == 0:
                        matrix_chk[loc_sqr[ii][0]][loc_sqr[ii][1]] = True
                    else:
                        matrix_chk[loc_sqr[ii][0]][loc_sqr[ii][1]] = False
            else:
                for ii in range(len(loc_sqr)):
                    matrix_chk[loc_sqr[ii][0]][loc_sqr[ii][1]] = False
count = 0
for i in range(1,n_str+1):
    for j in range(1, m_col+1):
        if matrix_chk[i][j]:
            count += 1
print(count)
