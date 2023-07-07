package com.exam.examportal.model.exam;

import com.exam.examportal.model.User;
import jakarta.persistence.*;

@Entity
public class Result {
    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private int result_id;

    @ManyToOne(fetch = FetchType.EAGER)
    private Quiz quiz;

    @ManyToOne(fetch = FetchType.EAGER)
    private User user;

    private int attempted;

    private int correctAnswers;

    private double marksGot;
    private int qAttempt;
    public int getqAttempt() {
        return qAttempt;
    }

    public void setqAttempt(int qAttempt) {
        this.qAttempt = qAttempt;
    }



    public int getResult_id() {
        return result_id;
    }

    public void setResult_id(int result_id) {
        this.result_id = result_id;
    }

    public Quiz getQuiz() {
        return quiz;
    }

    public void setQuiz(Quiz quiz) {
        this.quiz = quiz;
    }

    public User getUser() {
        return user;
    }

    public void setUser(User user) {
        this.user = user;
    }

    public int getAttempted() {
        return attempted;
    }

    public void setAttempted(int attempted) {
        this.attempted = attempted;
    }

    public int getCorrectAnswers() {
        return correctAnswers;
    }

    public void setCorrectAnswers(int correctAnswers) {
        this.correctAnswers = correctAnswers;
    }

    public double getMarksGot() {
        return marksGot;
    }

    public void setMarksGot(double marksGot) {
        this.marksGot = marksGot;
    }

    public Result(int result_id, Quiz quiz, User user, int attempted, int correctAnswers, double marksGot,int qAttempt) {
        super();
        this.result_id = result_id;
        this.quiz = quiz;
        this.user = user;
        this.attempted = attempted;
        this.correctAnswers = correctAnswers;
        this.marksGot = marksGot;
        this.qAttempt=qAttempt;
    }

    public Result() {
        super();
        // TODO Auto-generated constructor stub
    }

}
