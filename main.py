from fastapi import FastAPI
from pulp import LpProblem, LpMaximize, LpVariable, LpStatus, value

app = FastAPI()

@app.get("/solve_lp")
async def solve_linear_program():
    """
    แก้ปัญหา Linear Programming อย่างง่ายโดยใช้ PuLP

    ปัญหาตัวอย่าง:
    Maximize Z = 2x + 3y
    Subject to:
    x + y <= 10
    x >= 0
    y >= 0
    """
    # 1. สร้างปัญหา LP
    prob = LpProblem("Simple LP Problem", LpMaximize)

    # 2. สร้างตัวแปรตัดสินใจ
    x = LpVariable("x", 0, None)  # x >= 0, ไม่มีขีดจำกัดบน
    y = LpVariable("y", 0, None)  # y >= 0, ไม่มีขีดจำกัดบน

    # 3. กำหนดฟังก์ชันวัตถุประสงค์
    prob += 2 * x + 3 * y, "Objective Function"

    # 4. กำหนดข้อจำกัด
    prob += x + y <= 10, "Constraint 1: Sum of x and y"

    # 5. แก้ปัญหา
    prob.solve()

    # 6. ดึงผลลัพธ์
    status = LpStatus[prob.status]
    objective_value = value(prob.objective)
    x_value = x.varValue
    y_value = y.varValue

    return {
        "status": status,
        "objective_value": objective_value,
        "variables": {
            "x": x_value,
            "y": y_value
        },
        "message": "PuLP problem solved successfully."
    }
