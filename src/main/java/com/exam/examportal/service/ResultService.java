package com.exam.examportal.service;

import com.exam.examportal.model.User;
import com.exam.examportal.model.exam.Quiz;
import com.exam.examportal.model.exam.Result;

import java.util.List;

public interface ResultService {
    public Result addResult(Result result);
    public List<Result> getAllResult();
    public List<Result> getResultOfQuiz(Quiz quiz);
    public List<Result> getResultOfUser(User user);
    public List<Result> getResultOfUserAndQuiz(Quiz quiz,User user);
}
